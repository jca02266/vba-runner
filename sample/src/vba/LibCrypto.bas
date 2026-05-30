Option Explicit

' ============================================================
' LibCrypto.bas — 暗号化・ハッシング・エンコーディングユーティリティ
'
' Base64、Hex、URLエンコーディング関数を VBA で実装。
' Excel オブジェクトには依存せず単体テスト可能。
'
' 注: MD5、SHA256 などの暗号学的ハッシングは VBA での実装が
'     非常に複雑なため、スタブのままです。必要に応じて
'     外部ライブラリやカスタム実装を使用してください。
' ============================================================

' ============================================================
' Base64 エンコード・デコード
' ============================================================

' 文字列を Base64 形式にエンコードして返す
' 例: "Hello World" → "SGVsbG8gV29ybGQ="
Function Base64Encode(s As String) As String
    Const base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    Dim i As Long
    Dim j As Long
    Dim chunk(0 To 2) As Integer
    Dim result As String
    Dim padding As Integer
    Dim strLen As Long

    strLen = Len(s)
    If strLen = 0 Then
        Base64Encode = ""
        Exit Function
    End If

    i = 1
    Do While i <= strLen
        ' 3 バイトのチャンクを読み込む
        chunk(0) = Asc(Mid(s, i, 1))
        chunk(1) = IIf(i + 1 <= strLen, Asc(Mid(s, i + 1, 1)), 0)
        chunk(2) = IIf(i + 2 <= strLen, Asc(Mid(s, i + 2, 1)), 0)

        ' パディング判定
        If i + 2 <= strLen Then
            padding = 0
        ElseIf i + 1 <= strLen Then
            padding = 1
        Else
            padding = 2
        End If

        ' 24 ビットを 4 つの 6 ビット値に分割
        Dim b1, b2, b3, b4 As Integer
        b1 = (chunk(0) \ 4) + 1
        b2 = (((chunk(0) Mod 4) * 16) + (chunk(1) \ 16)) + 1
        b3 = (((chunk(1) Mod 16) * 4) + (chunk(2) \ 64)) + 1
        b4 = (chunk(2) Mod 64) + 1

        ' Base64 文字に変換
        result = result & Mid(base64Chars, b1, 1)
        result = result & Mid(base64Chars, b2, 1)

        If padding < 2 Then
            result = result & Mid(base64Chars, b3, 1)
        Else
            result = result & "="
        End If

        If padding < 1 Then
            result = result & Mid(base64Chars, b4, 1)
        Else
            result = result & "="
        End If

        i = i + 3
    Loop

    Base64Encode = result
End Function

' Base64 形式の文字列をデコードして返す
' 例: "SGVsbG8gV29ybGQ=" → "Hello World"
' 無効な Base64 形式の場合は Error 13 (Type Mismatch)
Function Base64Decode(s As String) As String
    Const base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    Dim result As String
    Dim i As Long
    Dim strLen As Long
    Dim pos As Long
    Dim b(0 To 3) As Integer  ' 4 つの 6 ビット値
    Dim groupCount As Integer
    Dim byte1, byte2, byte3 As Integer

    strLen = Len(s)
    If strLen = 0 Then
        Base64Decode = ""
        Exit Function
    End If

    groupCount = 0
    i = 1

    Do While i <= strLen
        Dim char As String
        char = Mid(s, i, 1)

        If char = "=" Then
            ' パディング文字 = はスキップ
            i = i + 1
        Else
            pos = InStr(base64Chars, char)
            If pos = 0 Then
                Err.Raise 13, "Base64Decode", "無効な Base64 文字: " & char
            End If

            ' 6 ビット値を保存（1～64 → 0～63）
            b(groupCount Mod 4) = pos - 1
            groupCount = groupCount + 1

            ' 4 文字グループが完成
            If groupCount Mod 4 = 0 Then
                ' 24 ビット → 3 バイト に変換
                byte1 = (b(0) * 4) + (b(1) \ 16)
                byte2 = ((b(1) Mod 16) * 16) + (b(2) \ 4)
                byte3 = ((b(2) Mod 4) * 64) + b(3)

                result = result & Chr(byte1) & Chr(byte2) & Chr(byte3)
            End If

            i = i + 1
        End If
    Loop

    ' 余りのバイトを処理（2 文字 または 3 文字）
    Dim remaining As Integer
    remaining = groupCount Mod 4

    If remaining = 2 Then
        ' 12 ビット → 1 バイト
        byte1 = (b(0) * 4) + (b(1) \ 16)
        result = result & Chr(byte1)
    ElseIf remaining = 3 Then
        ' 18 ビット → 2 バイト
        byte1 = (b(0) * 4) + (b(1) \ 16)
        byte2 = ((b(1) Mod 16) * 16) + (b(2) \ 4)
        result = result & Chr(byte1) & Chr(byte2)
    End If

    Base64Decode = result
End Function

' ============================================================
' MD5 ハッシング
' ============================================================

' 文字列の MD5 ハッシュ値を 16 進数文字列で返す
' 注: VBA で MD5 を完全実装するのは非常に複雑なため、スタブとします
' 必要な場合は外部 DLL（CAPICOM など）の使用を検討してください
Function MD5Hash(s As String) As String
    Err.Raise 5, "MD5Hash", "MD5Hash は VBA で実装されていません。外部ライブラリを使用してください"
End Function

' ============================================================
' SHA-256 ハッシング
' ============================================================

' 文字列の SHA-256 ハッシュ値を 16 進数文字列で返す
' 注: VBA で SHA256 を完全実装するのは非常に複雑なため、スタブとします
Function SHA256Hash(s As String) As String
    Err.Raise 5, "SHA256Hash", "SHA256Hash は VBA で実装されていません。外部ライブラリを使用してください"
End Function

' ============================================================
' SHA-1 ハッシング（参考: SHA256 の方が推奨）
' ============================================================

' 文字列の SHA-1 ハッシュ値を 16 進数文字列で返す
Function SHA1Hash(s As String) As String
    Err.Raise 5, "SHA1Hash", "SHA1Hash は VBA で実装されていません。外部ライブラリを使用してください"
End Function

' ============================================================
' HMAC（メッセージ認証コード）
' ============================================================

' 文字列を指定秘密鍵で HMAC-SHA256 署名して返す
Function HmacSHA256(message As String, secretKey As String) As String
    Err.Raise 5, "HmacSHA256", "HmacSHA256 は VBA で実装されていません。外部ライブラリを使用してください"
End Function

' 文字列を指定秘密鍵で HMAC-SHA1 署名して返す
Function HmacSHA1(message As String, secretKey As String) As String
    Err.Raise 5, "HmacSHA1", "HmacSHA1 は VBA で実装されていません。外部ライブラリを使用してください"
End Function

' ============================================================
' 16 進数エンコード・デコード
' ============================================================

' 文字列を 16 進数文字列にエンコードして返す
' 例: "AB" → "4142"（'A' = 0x41, 'B' = 0x42）
Function HexEncode(s As String) As String
    Dim result As String
    Dim i As Long
    Dim charCode As Integer

    For i = 1 To Len(s)
        charCode = Asc(Mid(s, i, 1))
        result = result & Right("0" & Hex(charCode), 2)
    Next i

    HexEncode = result
End Function

' 16 進数文字列をデコードして返す
' 例: "4142" → "AB"
Function HexDecode(hexStr As String) As String
    Dim result As String
    Dim i As Long
    Dim hexPair As String
    Dim byteValue As Integer

    ' 奇数文字の場合はエラー
    If Len(hexStr) Mod 2 <> 0 Then
        Err.Raise 13, "HexDecode", "16進数文字列の長さは偶数である必要があります"
    End If

    For i = 1 To Len(hexStr) Step 2
        hexPair = Mid(hexStr, i, 2)
        ' 16進数文字列の妥当性チェック
        If Not IsHexString(hexPair) Then
            Err.Raise 13, "HexDecode", "無効な16進数文字が含まれています: " & hexPair
        End If
        ' 16進数を手動で10進数に変換
        byteValue = HexCharToInt(Left(hexPair, 1)) * 16 + HexCharToInt(Right(hexPair, 1))
        result = result & Chr(byteValue)
    Next i

    HexDecode = result
End Function

' 16進数文字を10進数に変換（0-F → 0-15）
Private Function HexCharToInt(c As String) As Integer
    Dim charCode As Integer
    c = UCase(c)
    charCode = Asc(c)

    If charCode >= Asc("0") And charCode <= Asc("9") Then
        HexCharToInt = charCode - Asc("0")
    ElseIf charCode >= Asc("A") And charCode <= Asc("F") Then
        HexCharToInt = charCode - Asc("A") + 10
    Else
        HexCharToInt = 0
    End If
End Function

' 文字列がすべて有効な 16 進数文字（0-9, A-F）かどうかを返す
Private Function IsHexString(s As String) As Boolean
    Dim i As Long
    Dim c As String

    IsHexString = True  ' 初期値を True に設定

    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If Not ((c >= "0" And c <= "9") Or (UCase(c) >= "A" And UCase(c) <= "F")) Then
            IsHexString = False
            Exit Function
        End If
    Next i
End Function

' ============================================================
' URL エンコード補助（RFC 3986 準拠）
' ============================================================

' 文字列を URL セーフな形式にエンコードして返す
' 予約文字（! * ' ( ) ; : @ & = + $ , / ? # [ ]）は % エンコード
' セーフ文字（A-Z a-z 0-9 - _ . ~）はそのまま
Function URLEncodeRFC3986(s As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String
    Dim charCode As Integer

    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        charCode = Asc(c)

        ' セーフ文字: 0-9 (48-57), A-Z (65-90), a-z (97-122), - _ . ~
        If (charCode >= 48 And charCode <= 57) Or _
           (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Or _
           c = "-" Or c = "_" Or c = "." Or c = "~" Then
            result = result & c
        Else
            ' パーセントエンコード
            result = result & "%" & Right("0" & Hex(charCode), 2)
        End If
    Next i

    URLEncodeRFC3986 = result
End Function
