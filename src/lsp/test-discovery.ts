import {
    ClassDeclaration,
    ProcedureDeclaration,
    Statement,
} from '../engine/parser';

export interface TestItem {
    id: string;
    label: string;
    name: string;
    type: 'Test' | 'Suite';
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
    children?: TestItem[];
}

export class TestDiscovery {
    discoverTests(statements: Statement[]): TestItem[] {
        const tests: TestItem[] = [];

        for (const stmt of statements) {
            const discovered = this.discoverInStatement(stmt);
            if (discovered.length > 0) {
                tests.push(...discovered);
            }
        }

        return tests;
    }

    private discoverInStatement(stmt: Statement): TestItem[] {
        if (stmt.type === 'ProcedureDeclaration') {
            return this.discoverTestProcedure(stmt as ProcedureDeclaration);
        } else if (stmt.type === 'ClassDeclaration') {
            return this.discoverTestClass(stmt as ClassDeclaration);
        }

        return [];
    }

    private discoverTestProcedure(proc: ProcedureDeclaration): TestItem[] {
        // Check if procedure name starts with "Test_" (case-insensitive)
        const name = proc.name.name;
        if (!name.toLowerCase().startsWith('test_')) {
            return [];
        }

        // Create test item
        const id = this.generateId(name);
        const range = proc.loc
            ? {
                start: { line: proc.loc.start.line - 1, character: proc.loc.start.column - 1 },
                end: { line: proc.loc.end.line - 1, character: proc.loc.end.column - 1 },
            }
            : {
                start: { line: 0, character: 0 },
                end: { line: 0, character: 0 },
            };

        return [
            {
                id,
                label: name,
                name,
                type: 'Test',
                range,
            },
        ];
    }

    private discoverTestClass(cls: ClassDeclaration): TestItem[] {
        const tests: TestItem[] = [];
        const children: TestItem[] = [];

        // Look for Test_ methods inside the class
        for (const stmt of cls.body) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                const procTests = this.discoverTestProcedure(proc);
                children.push(...procTests);
            }
        }

        // If the class has test methods, create a suite item
        if (children.length > 0) {
            tests.push({
                id: this.generateId(cls.name),
                label: cls.name,
                name: cls.name,
                type: 'Suite',
                range: {
                    start: { line: 0, character: 0 },
                    end: { line: 0, character: 0 },
                },
                children,
            });
        }

        return tests;
    }

    private generateId(name: string): string {
        // Generate a unique ID from the name
        // Simple approach: use the name itself
        // In a real implementation, would include file path, line number, etc.
        return name.toLowerCase().replace(/[^a-z0-9_]/g, '_');
    }
}
