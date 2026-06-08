/**
 * Implements Directive (§5.2.4.2) のテスト
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

// --- 1. Implements のパースと実行 (No-op) ---
const implementsCode = `
    Class IAnimal
        Public Sub Speak()
        End Sub
    End Class

    Class Dog
        Implements IAnimal
        
        Public Sub IAnimal_Speak()
            Debug.Print "Woof"
        End Sub
    End Class

    Sub Test()
        Dim d As Dog
        Set d = New Dog
        d.IAnimal_Speak
    End Sub
`;

const ev1 = evalVBA(implementsCode);
ev1.callProcedure('Test', []);
console.log('[PASS] Implements ディレクティブ');

console.log('\n✅ Implements Directive: 全テスト通過');
