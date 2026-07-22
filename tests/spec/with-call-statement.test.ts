import { evalVBAModules, assert } from '../../test-libs/test-runner';

const ev = evalVBAModules([
    {
        name: 'Basket',
        parseAsClass: 'Basket',
        code: `
            Private mItems As Collection

            Private Sub Class_Initialize()
                Set mItems = New Collection
            End Sub

            Public Sub AddItem(ByVal value As String)
                mItems.Add value
            End Sub

            Public Property Get Count() As Long
                Count = mItems.Count
            End Property
        `,
    },
    {
        name: 'Main',
        code: `
            Function TestWithCall() As Long
                Dim basket As New Basket
                With basket
                    Call .AddItem "A"
                    Call .AddItem "B"
                End With
                TestWithCall = basket.Count
            End Function
        `,
    },
]);

assert.strictEqual(ev.callProcedure('TestWithCall', []), 2);

console.log('✅ Call .Method arg works inside With blocks');
