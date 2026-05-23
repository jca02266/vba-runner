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
        html, body {
            width: 100%;
            height: 100%;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: var(--vscode-editor-background, #1e1e1e);
            color: var(--vscode-editor-foreground, #d4d4d4);
        }
        body {
            display: flex;
            flex-direction: column;
            padding: 12px;
            overflow: hidden;
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            flex-wrap: wrap;
            gap: 12px;
            flex-shrink: 0;
        }
        h1 { font-size: 16px; font-weight: 600; }
        .info { font-size: 12px; opacity: 0.8; }
        .buttons {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            align-items: center;
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
            white-space: nowrap;
        }
        button:hover {
            background: var(--vscode-button-hoverBackground, #1177bb);
        }
        .zoom-controls {
            display: flex;
            gap: 4px;
            align-items: center;
            border-left: 1px solid var(--vscode-editor-lineHighlightBorderBackground, #3e3e42);
            padding-left: 8px;
            margin-left: 8px;
        }
        .zoom-level {
            font-size: 11px;
            min-width: 40px;
            text-align: center;
            color: var(--vscode-editor-foreground, #d4d4d4);
        }
        .mermaid-container {
            flex: 1;
            overflow: auto;
            border: 1px solid var(--vscode-editor-lineHighlightBorderBackground, #3e3e42);
            border-radius: 4px;
            padding: 8px;
            background: var(--vscode-editor-background, #1e1e1e);
            min-height: 200px;
        }
        .mermaid-container > .mermaid {
            display: flex;
            justify-content: center;
            align-items: flex-start;
            min-width: min-content;
            padding: 16px;
        }
        .mermaid svg {
            max-width: 100%;
            height: auto;
        }
        .mermaid g.node text {
            font-size: 14px !important;
            font-weight: 500;
        }
        .mermaid g.edgeLabel text {
            font-size: 12px !important;
        }
        .legend {
            margin-top: 8px;
            font-size: 11px;
            padding: 8px;
            border-top: 1px solid var(--vscode-editor-lineHighlightBorderBackground, #3e3e42);
            flex-shrink: 0;
        }
        .legend-item {
            display: inline-block;
            margin-right: 16px;
            margin-top: 4px;
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
            <div class="zoom-controls">
                <button onclick="zoomOut()" style="padding: 4px 8px;">−</button>
                <div class="zoom-level" id="zoomLevel">100%</div>
                <button onclick="zoomIn()" style="padding: 4px 8px;">+</button>
                <button onclick="resetZoom()" style="padding: 4px 8px;">Reset</button>
            </div>
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
        let currentZoom = 1.0;
        const minZoom = 0.2;
        const maxZoom = 20.0;
        const zoomStep = 0.5;

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

        function updateZoom(zoomFactor) {
            currentZoom = Math.max(minZoom, Math.min(maxZoom, zoomFactor));
            const svg = document.querySelector('.mermaid svg');
            if (svg) {
                svg.style.transform = 'scale(' + currentZoom + ')';
                svg.style.transformOrigin = 'top center';
                svg.style.transformBox = 'fill-box';
            }
            document.getElementById('zoomLevel').textContent = Math.round(currentZoom * 100) + '%';
        }

        function zoomIn() {
            updateZoom(currentZoom + zoomStep);
        }

        function zoomOut() {
            updateZoom(currentZoom - zoomStep);
        }

        function resetZoom() {
            updateZoom(1.0);
        }

        // Initialize Mermaid with optimized layout
        mermaid.initialize({
            startOnLoad: true,
            theme: 'dark',
            securityLevel: 'loose',
            flowchart: {
                curve: 'linear',
                nodeSpacing: 40,
                rankSpacing: 60,
                padding: 16,
                useMaxWidth: false,
                htmlLabels: true,
                fontSize: 14
            },
            fontFamily: 'inherit'
        });
        mermaid.contentLoaded();

        // マウスホイールズーム（Ctrl キー押下時）
        document.addEventListener('wheel', (e) => {
            if (e.ctrlKey || e.metaKey) {
                e.preventDefault();
                if (e.deltaY < 0) {
                    zoomIn();
                } else {
                    zoomOut();
                }
            }
        }, { passive: false });

        // Mermaid レンダリング後にズーム適用
        setTimeout(() => {
            const svg = document.querySelector('.mermaid svg');
            if (svg) {
                updateZoom(1.0);
            }
        }, 100);
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
