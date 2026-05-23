import { CallGraph } from './call-graph-provider';
import * as vscode from 'vscode';
import * as crypto from 'crypto';

export function generateCallGraphHtml(
    graph: CallGraph,
    _webview: vscode.Webview,
    _extensionUri: vscode.Uri,
    focusProcName: string | null
): string {
    const nonce = crypto.randomBytes(16).toString('base64');

    // Filter nodes and edges if focusProcName is provided
    let visibleNodes = new Set<string>();
    let visibleEdges = graph.edges;

    if (focusProcName) {
        const focusKey = focusProcName.toLowerCase();
        if (graph.nodes.has(focusKey)) {
            visibleNodes = getTransitiveNodes(graph, focusKey);
        } else {
            visibleNodes = new Set(graph.nodes.keys());
        }
        visibleEdges = graph.edges.filter(
            (e) => visibleNodes.has(e.from) && visibleNodes.has(e.to)
        );
    } else {
        visibleNodes = new Set(graph.nodes.keys());
    }

    const mermaidText = buildMermaidText(graph, visibleNodes, visibleEdges);

    return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VBA Call Graph</title>
    <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: var(--vscode-editor-background, #1e1e1e);
            color: var(--vscode-editor-foreground, #d4d4d4);
            padding: 16px;
            max-height: 100vh;
            overflow: hidden;
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            flex-wrap: wrap;
            gap: 12px;
        }
        h1 { font-size: 18px; font-weight: 500; }
        .info { font-size: 13px; opacity: 0.8; }
        .buttons {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        button {
            padding: 6px 12px;
            background: var(--vscode-button-background, #0e639c);
            color: var(--vscode-button-foreground, #fff);
            border: none;
            border-radius: 2px;
            cursor: pointer;
            font-size: 12px;
            transition: background 0.2s;
        }
        button:hover {
            background: var(--vscode-button-hoverBackground, #1177bb);
        }
        .mermaid-container {
            width: 100%;
            height: calc(100vh - 120px);
            overflow: auto;
            border: 1px solid var(--vscode-editor-lineHighlightBorderBackground, #3e3e42);
            border-radius: 4px;
            padding: 12px;
            background: var(--vscode-editor-background, #1e1e1e);
        }
        .mermaid {
            min-height: 100%;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .legend {
            margin-top: 12px;
            font-size: 12px;
            padding: 8px;
            border-top: 1px solid var(--vscode-editor-lineHighlightBorderBackground, #3e3e42);
        }
        .legend-item {
            display: inline-block;
            margin-right: 16px;
            margin-top: 6px;
        }
        .legend-color {
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 2px;
            margin-right: 4px;
            vertical-align: middle;
        }
        .color-entry { background: #6699ff; }
        .color-dead { background: #ff8888; }
        .color-excel { border: 2px solid #ff9900; }
    </style>
</head>
<body>
    <div class="header">
        <div>
            <h1>VBA Call Graph</h1>
            <p class="info">${visibleNodes.size} procedures, ${visibleEdges.length} calls${focusProcName ? ` (focused on ${focusProcName})` : ''}</p>
        </div>
        <div class="buttons">
            <button onclick="copyAsMarkdown()">📋 Copy as Mermaid</button>
            ${focusProcName ? `<button onclick="showFullGraph()">🔄 Show Full Graph</button>` : ''}
        </div>
    </div>

    <div class="mermaid-container">
        <div class="mermaid">
${mermaidText}
        </div>
    </div>

    <div class="legend">
        <strong>Legend:</strong>
        <div class="legend-item">
            <span class="legend-color color-entry"></span> Entry point (Public, 0 refs)
        </div>
        <div class="legend-item">
            <span class="legend-color color-dead"></span> Dead code (Private, 0 refs)
        </div>
        <div class="legend-item">
            <span class="legend-color color-excel"></span> Excel-dependent
        </div>
    </div>

    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const mermaidSource = ${JSON.stringify(mermaidText)};

        function nodeClicked(nodeId) {
            const cleanId = nodeId.replace(/["']/g, '');
            vscode.postMessage({ type: 'goToDefinition', procName: cleanId });
        }

        function copyAsMarkdown() {
            navigator.clipboard.writeText(mermaidSource).then(() => {
                alert('Mermaid source copied to clipboard');
            }).catch(() => {
                alert('Failed to copy');
            });
        }

        function showFullGraph() {
            vscode.postMessage({ type: 'showFullGraph' });
        }

        // Initialize Mermaid
        mermaid.initialize({ startOnLoad: true, theme: 'dark' });
        mermaid.contentLoaded();
    </script>
</body>
</html>`;
}

function buildMermaidText(graph: CallGraph, visibleNodes: Set<string>, visibleEdges: any[]): string {
    const lines: string[] = ['graph TD'];

    // Define classDefs
    lines.push('    classDef entry fill:#6699ff,stroke:#3366cc,color:#fff');
    lines.push('    classDef dead fill:#ff8888,stroke:#cc3333,color:#fff');
    lines.push('    classDef excel stroke:#ff9900,stroke-width:3px');
    lines.push('    classDef deadExcel fill:#ff8888,stroke:#ff9900,stroke-width:3px,color:#fff');

    // Add visible nodes with appropriate classes
    for (const key of visibleNodes) {
        const node = graph.nodes.get(key);
        if (!node) continue;

        // Use node name as-is (Mermaid handles identifiers)
        const nodeId = node.name;
        let classDef = '';

        const isEntryPoint = node.scope === 'public' && node.referenceCount === 0;
        const isDeadCode = node.scope === 'private' && node.referenceCount === 0;
        const isExcelDep = node.isExcelDependent;

        if (isDeadCode && isExcelDep) {
            classDef = ':::deadExcel';
        } else if (isEntryPoint) {
            classDef = ':::entry';
        } else if (isDeadCode) {
            classDef = ':::dead';
        } else if (isExcelDep) {
            classDef = ':::excel';
        }

        if (classDef) {
            lines.push(`    ${nodeId}${classDef}`);
        } else {
            lines.push(`    ${nodeId}`);
        }

        // Add click handler using callback syntax
        lines.push(`    click ${nodeId} nodeClicked`);
    }

    // Add edges
    for (const edge of visibleEdges) {
        const fromName = graph.nodes.get(edge.from)?.name || '';
        const toName = graph.nodes.get(edge.to)?.name || '';
        lines.push(`    ${fromName} --> ${toName}`);
    }

    return lines.join('\n');
}

function getTransitiveNodes(graph: CallGraph, focusKey: string): Set<string> {
    const visited = new Set<string>();
    const result = new Set<string>([focusKey]);

    // BFS for ancestors and descendants
    const queue = [focusKey];

    while (queue.length > 0) {
        const current = queue.shift()!;
        if (visited.has(current)) continue;
        visited.add(current);

        // Find descendants (current -> X)
        for (const edge of graph.edges) {
            if (edge.from === current && !result.has(edge.to)) {
                result.add(edge.to);
                queue.push(edge.to);
            }
        }

        // Find ancestors (X -> current)
        for (const edge of graph.edges) {
            if (edge.to === current && !result.has(edge.from)) {
                result.add(edge.from);
                queue.push(edge.from);
            }
        }
    }

    return result;
}
