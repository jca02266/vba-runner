import * as vscode from 'vscode';
import { DebugAdapter } from './debug-adapter';
import { LSPServer } from './server';
import * as fs from 'fs';

export class VBADebugAdapterFactory implements vscode.DebugAdapterDescriptorFactory {
    constructor(private lspServer: LSPServer) {}

    createDebugAdapterDescriptor(session: vscode.DebugSession): vscode.ProviderResult<vscode.DebugAdapterDescriptor> {
        const program = session.configuration.program as string | undefined;
        if (!program) return undefined;

        const uri = vscode.Uri.file(program).toString();

        // Ensure the document is loaded in LSPServer
        if (!this.lspServer.hasDocument(uri)) {
            try {
                const content = fs.readFileSync(program, 'utf-8');
                this.lspServer.didOpen(uri, content);
            } catch {
                return undefined;
            }
        }

        const adapter = this.lspServer.createDebugAdapter(uri);
        if (!adapter) return undefined;

        return new vscode.DebugAdapterInlineImplementation(new VBAInlineDebugAdapter(adapter));
    }
}

class VBAInlineDebugAdapter implements vscode.DebugAdapter {
    private _emitter = new vscode.EventEmitter<vscode.DebugProtocolMessage>();
    readonly onDidSendMessage: vscode.Event<vscode.DebugProtocolMessage> = this._emitter.event;
    private _seq = 1;

    constructor(private adapter: DebugAdapter) {}

    handleMessage(message: vscode.DebugProtocolMessage): void {
        const msg = message as any;
        if (msg.type !== 'request') return;

        const result = this.adapter.handleRequest({ command: msg.command, arguments: msg.arguments });
        const success = result == null || result.success !== false;

        this._send({
            type: 'response',
            request_seq: msg.seq,
            seq: this._seq++,
            success,
            command: msg.command,
            body: success ? result : undefined,
            message: success ? undefined : result?.error,
        });

        // Synthetic events required by the DAP handshake
        if (msg.command === 'initialize') {
            this._send({ type: 'event', seq: this._seq++, event: 'initialized' });
        } else if (msg.command === 'launch' || msg.command === 'configurationDone') {
            this._send({ type: 'event', seq: this._seq++, event: 'stopped', body: { reason: 'entry', threadId: 1, allThreadsStopped: true } });
        } else if (msg.command === 'continue' || msg.command === 'next' || msg.command === 'stepIn' || msg.command === 'stepOut') {
            this._send({ type: 'event', seq: this._seq++, event: 'stopped', body: { reason: 'step', threadId: 1, allThreadsStopped: true } });
        } else if (msg.command === 'disconnect') {
            this._send({ type: 'event', seq: this._seq++, event: 'terminated' });
        }
    }

    private _send(msg: any): void {
        this._emitter.fire(msg);
    }

    dispose(): void {
        this._emitter.dispose();
    }
}
