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

// Bug CF: `TypeOf obj Is InterfaceName` was returning False when obj's class Implements the interface.
// Only __vbaTypeName__ (the concrete class name) was checked, not the Implements chain.
{
    const code = `
Class Animal
    Public Function Speak() As String : End Function
End Class
Class Dog
    Implements Animal
    Public Function Animal_Speak() As String
        Animal_Speak = "Woof"
    End Function
End Class
Function TestCF_TypeOf() As String
    Dim d As New Dog
    Dim a As Animal
    Set a = d
    Dim r As String
    If TypeOf a Is Animal Then r = "IsAnimal" Else r = "NotAnimal"
    If TypeOf a Is Dog Then r = r & ",IsDog" Else r = r & ",NotDog"
    If TypeOf d Is Animal Then r = r & ",DogIsAnimal" Else r = r & ",DogNotAnimal"
    TestCF_TypeOf = r
End Function
`;
    const ev2 = evalVBA(code);
    assert.strictEqual(ev2.callProcedure('TestCF_TypeOf', []), 'IsAnimal,IsDog,DogIsAnimal',
        'Bug CF: TypeOf obj Is InterfaceName returns True when obj implements the interface');
    console.log('[PASS] Bug CF: TypeOf obj Is InterfaceName — Implements チェーンを参照');
}

console.log('\n✅ Implements Directive: 全テスト通過');
