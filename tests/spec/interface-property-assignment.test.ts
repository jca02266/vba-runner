/**
 * Implements interface Property Let/Set assignments must dispatch to the
 * InterfaceName_Property implementation for both direct and indexed targets.
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string): any {
    return evalVBASingle(code).callProcedure(name, []);
}

const code = String.raw`
Class Payload
    Public Name As String
End Class

Class IStore
    Public Property Set Item(ByVal index As Long, ByVal value As Object)
    End Property
    Public Property Set Value(ByVal value As Object)
    End Property
End Class

Class Store
    Implements IStore
    Public Seen As String
    Public Property Set IStore_Item(ByVal index As Long, ByVal value As Object)
        Seen = CStr(index) & ":" & value.Name
    End Property
    Public Property Set IStore_Value(ByVal value As Object)
        Seen = "0:" & value.Name
    End Property
End Class

Class ILabel
    Public Property Let Label(ByVal index As Long, ByVal value As String)
    End Property
    Public Property Let Text(ByVal value As String)
    End Property
End Class

Class Labels
    Implements ILabel
    Public Seen As String
    Public Property Let ILabel_Label(ByVal index As Long, ByVal value As String)
        Seen = CStr(index) & ":" & value
    End Property
    Public Property Let ILabel_Text(ByVal value As String)
        Seen = "0:" & value
    End Property
End Class

Function TestInterfaceSet() As String
    Dim concrete As New Store, target As IStore, payload As New Payload
    payload.Name = "object"
    Set target = concrete
    Set target.Item(2) = payload
    TestInterfaceSet = concrete.Seen
End Function

Function TestInterfaceLet() As String
    Dim concrete As New Labels, target As ILabel
    Set target = concrete
    target.Label(4) = "text"
    TestInterfaceLet = concrete.Seen
End Function

Function TestInterfaceWithSet() As String
    Dim concrete As New Store, target As IStore, payload As New Payload
    payload.Name = "with"
    Set target = concrete
    With target
        Set .Item(3) = payload
    End With
    TestInterfaceWithSet = concrete.Seen
End Function

Function TestInterfaceUnindexedSet() As String
    Dim concrete As New Store, target As IStore, payload As New Payload
    payload.Name = "plain"
    Set target = concrete
    Set target.Value = payload
    TestInterfaceUnindexedSet = concrete.Seen
End Function

Function TestInterfaceUnindexedLet() As String
    Dim concrete As New Labels, target As ILabel
    Set target = concrete
    target.Text = "plain"
    TestInterfaceUnindexedLet = concrete.Seen
End Function
`;

assert.strictEqual(runFunc(code, 'TestInterfaceSet'), '2:object',
    'indexed Interface Property Set dispatches to the implementation');
assert.strictEqual(runFunc(code, 'TestInterfaceLet'), '4:text',
    'indexed Interface Property Let dispatches to the implementation');
assert.strictEqual(runFunc(code, 'TestInterfaceWithSet'), '3:with',
    'With Interface Property Set dispatches to the implementation');
assert.strictEqual(runFunc(code, 'TestInterfaceUnindexedSet'), '0:plain',
    'unindexed Interface Property Set dispatches to the implementation');
assert.strictEqual(runFunc(code, 'TestInterfaceUnindexedLet'), '0:plain',
    'unindexed Interface Property Let dispatches to the implementation');
console.log('[PASS] Implements interface Property Let/Set assignment dispatch');
