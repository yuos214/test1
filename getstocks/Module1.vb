Imports System.Text
Imports System.Net
Module Module1
    'Const TextPath1 As String = "D:\stocks\Avelist.txt"
    'Const TextPath1 As String = "D:\stocks\errorlist.txt"
    Const Url1 As String = "http://stocks.finance.yahoo.co.jp/stocks/detail/?code="
    Const Url2 As String = "http://info.finance.yahoo.co.jp/history/?code="
    Const Url3 As String = "http://finance.yahoo.com/q/hp?s=%5EDJI"

    Sub Start()

        Dim FolderName1 As String
        FolderName1 = Form1.TextBoxWorkFolderName1.Text

        Dim ListFileName1 As String
        ListFileName1 = Form1.TextBoxListFileName1.Text

        Dim TextPath1 As String
        TextPath1 = FolderName1 & "\" & ListFileName1

        Dim NumCollection1 As Collection
        NumCollection1 = New Collection

        NumCollection1 = GetCodeList(TextPath1)

        Dim HtmlString1 As String

        Dim ParamCollection1 As Collection
        ParamCollection1 = New Collection

        Dim i As Long
        For i = 1 To NumCollection1.Count

            HtmlString1 = GetHtml(NumCollection1.Item(i))

            Form1.TextBox1.Text &= i & " / " & NumCollection1.Count

            Dim ParamString1 As String
            ParamString1 = NumCollection1.Item(i) & " " & GetDetail(HtmlString1, "PER") & " " & GetDetail(HtmlString1, "PBR") & " " & GetDetail(HtmlString1, "最低購入代金", True)

            ParamCollection1.Add(ParamString1)

            Form1.TextBox1.Text &= " - " & ParamString1 & Environment.NewLine

            'カレット位置を末尾に移動
            Form1.TextBox1.SelectionStart = Form1.TextBox1.Text.Length
            'テキストボックスにフォーカスを移動
            Form1.TextBox1.Focus()
            'カレット位置までスクロール
            Form1.TextBox1.ScrollToCaret()

            Form1.Refresh()

        Next

        Dim textFile As System.IO.StreamWriter

        textFile = New System.IO.StreamWriter(FolderName1 & "\Parameters.txt")
        For i = 1 To ParamCollection1.Count
            textFile.WriteLine(ParamCollection1.Item(i))
        Next
        textFile.Close()

        MsgBox("計算終了")

    End Sub


    Sub Start2(Optional ByVal iUSAFlag1 As Boolean = False)

        Dim FolderName1 As String
        FolderName1 = Form1.TextBoxWorkFolderName1.Text

        Dim ListFileName1 As String
        ListFileName1 = Form1.TextBoxListFileName1.Text

        Dim TextPath1 As String
        TextPath1 = FolderName1 & "\" & ListFileName1

        Dim NumCollection1 As Collection
        NumCollection1 = New Collection

        If iUSAFlag1 = False Then
            NumCollection1 = GetCodeList(TextPath1)
        Else
            NumCollection1.Add("EDJI")
        End If

        'Dim HtmlString1 As String

        Dim ParamCollection1 As Collection
        ParamCollection1 = New Collection

        Dim i, j, k As Long
        For i = 1 To NumCollection1.Count

            'HtmlString1 = GetHtml(NumCollection1.Item(i))

            'Form1.TextBox1.Text &= i & " / " & NumCollection1.Count
            Form1.TextBox2.Text = i & " / " & NumCollection1.Count

            System.Threading.Thread.Sleep(5000)

            'Dim ParamString1 As String
            'ParamString1 = NumCollection1.Item(i) & " " & GetDetail(HtmlString1, "PER") & " " & GetDetail(HtmlString1, "PBR") & " " & GetDetail(HtmlString1, "最低購入代金", True)

            'ParamCollection1.Add(ParamString1)

            'Form1.TextBox1.Text &= " - " & ParamString1 & Environment.NewLine

            Dim colHtml1 As Collection
            colHtml1 = GetHistoryHtml(NumCollection1.Item(i), DateTime.Parse("2005/01/01"), System.DateTime.Today, iUSAFlag1)

            Dim colHiostory1 As Collection
            colHiostory1 = New Collection

            For j = 1 To colHtml1.Count

                Dim tmpText2 As String
                tmpText2 = colHtml1.Item(j)

                If tmpText2.IndexOf("該当する期間のデータはありません。") = -1 Then

                    Dim tmpCollection1 As Collection
                    tmpCollection1 = GetHistory(tmpText2, iUSAFlag1)
                    For k = tmpCollection1.Count To 1 Step -1
                        colHiostory1.Add(tmpCollection1.Item(k))
                    Next

                End If

            Next

            Dim tmpText1 As String
            tmpText1 = NumCollection1.Item(i)

            Dim textFile As System.IO.StreamWriter

            textFile = New System.IO.StreamWriter(FolderName1 & "\History\" & tmpText1 & ".txt")
            For k = 1 To colHiostory1.Count
                textFile.WriteLine(colHiostory1.Item(k))
            Next
            textFile.Close()

            'Form1.TextBox1.Text &= Environment.NewLine

            ''カレット位置を末尾に移動
            'Form1.TextBox1.SelectionStart = Form1.TextBox1.Text.Length
            ''テキストボックスにフォーカスを移動
            'Form1.TextBox1.Focus()
            ''カレット位置までスクロール
            'Form1.TextBox1.ScrollToCaret()

            Form1.Refresh()

        Next

        Form1.TextBox2.Text = "計算終了"
        MsgBox("計算終了")

    End Sub

    Function GetHtml(ByVal iCode As String) As String
        GetHtml = ""

        Dim strUrl As String 'URL

        Dim myWebClient As New WebClient    'Web
        Dim myDatabuffer As Byte()          'バッファバイト配列

        Dim strHtml As String

        'URLをセットします
        strUrl = Url1 & iCode

        'URLからデータを取り出します
        System.Threading.Thread.Sleep(1000)
        myDatabuffer = myWebClient.DownloadData(strUrl)
        'エンコードしています
        strHtml = Encoding.UTF8.GetString(myDatabuffer)

        'ページソースを表示しています
        GetHtml = strHtml
    End Function

    Function GetHistoryHtml(ByVal iCode As String, ByVal iSDate As Date, ByVal iEDate As Date, Optional ByVal iUSAFlag As Boolean = False) As Collection

        GetHistoryHtml = New Collection

        Dim EndFlag1 As Boolean = False

        Dim SDate1, EDate1 As Date
        SDate1 = iSDate
        EDate1 = iSDate.AddDays(28)

        Do

            Form1.TextBox1.Text = SDate1 & " -> " & EDate1 & Environment.NewLine

            If EDate1 > iEDate Then
                EDate1 = iEDate
                EndFlag1 = True
            End If

            Dim SDateArray1 As Long()
            ReDim SDateArray1(2)
            SDateArray1(0) = SDate1.Year
            SDateArray1(1) = SDate1.Month
            SDateArray1(2) = SDate1.Day

            Dim EDateArray1 As Long()
            ReDim EDateArray1(2)
            EDateArray1(0) = EDate1.Year
            EDateArray1(1) = EDate1.Month
            EDateArray1(2) = EDate1.Day

            Dim strUrl As String 'URL

            Dim myWebClient As New WebClient    'Web
            Dim myDatabuffer As Byte()          'バッファバイト配列

            Dim strHtml As String

            'URLをセットします
            If iUSAFlag = False Then
                strUrl = Url2 & iCode & "&sy=" & SDateArray1(0) & "&sm=" & SDateArray1(1) & "&sd=" & SDateArray1(2) & "&ey=" & EDateArray1(0) & "&em=" & EDateArray1(1) & "&ed=" & EDateArray1(2) & "&tm=d"
            Else
                strUrl = Url3 & "&a=" & SDateArray1(1) - 1 & "&b=" & SDateArray1(2) & "&c=" & SDateArray1(0) & "&d=" & EDateArray1(1) - 1 & "&e=" & EDateArray1(2) & "&f=" & EDateArray1(0) & "&g=d"
            End If

            Dim ErrCnt1 As Long
            ErrCnt1 = 0

Retry1:
            If ErrCnt1 >= 5 Then Exit Function

            'URLからデータを取り出します
            Application.DoEvents()
            System.Threading.Thread.Sleep(1000)
            Try
                myDatabuffer = myWebClient.DownloadData(strUrl)
            Catch
                ErrCnt1 += 1
                Form1.TextBox1.Text &= "リトライ" & ErrCnt1 & "回目" & Environment.NewLine
                System.Threading.Thread.Sleep(300000)
                GoTo Retry1
            End Try
            'エンコードしています
            strHtml = Encoding.UTF8.GetString(myDatabuffer)

            GetHistoryHtml.Add(strHtml)

            SDate1 = EDate1.AddDays(1)
            EDate1 = SDate1.AddDays(28)

        Loop Until EndFlag1



        'カレット位置を末尾に移動
        'Form1.TextBox1.SelectionStart = Form1.TextBox1.Text.Length
        ' ''テキストボックスにフォーカスを移動
        ''Form1.TextBox1.Focus()
        ''カレット位置までスクロール
        'Form1.TextBox1.ScrollToCaret()

    End Function

    Function GetCodeList(ByVal iPath As String) As Collection

        GetCodeList = New Collection

        ' StreamReader の新しいインスタンスを生成する
        Dim cReader As New System.IO.StreamReader(iPath, System.Text.Encoding.Default)

        ' 読み込んだ結果をすべて格納するための変数を宣言する
        Dim stResult As String = String.Empty

        ' 読み込みできる文字がなくなるまで繰り返す
        While (cReader.Peek() >= 0)
            ' ファイルを 1 行ずつ読み込む
            Dim stBuffer As String = cReader.ReadLine()
            ' 読み込んだものを追加で格納する
            GetCodeList.Add(stBuffer)
        End While

        ' cReader を閉じる (正しくは オブジェクトの破棄を保証する を参照)
        cReader.Close()

    End Function

    Function GetDetail(ByVal iString As String, ByVal iSearch As String, Optional ByVal Comma As Boolean = False) As String
        GetDetail = ""

        Dim r As New System.Text.RegularExpressions.Regex( _
        ".*<strong>.*\n.*\b" & iSearch & "\b", _
        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        'TextBox1.Text内で正規表現と一致する対象を1つ検索 
        Dim m As System.Text.RegularExpressions.Match = r.Match(iString)

        '次のように一致する対象をすべて検索することもできる 
        'Dim mc As System.Text.RegularExpressions.MatchCollection = _
        '    r.Matches(TextBox1.Text)

        'While m.Success
        '一致した対象が見つかったときキャプチャした部分文字列を表示 
        Dim Text1
        Text1 = m.Value
        '次に一致する対象を検索 
        'm = m.NextMatch()
        'End While

        Dim Separater1 As String
        If Comma = False Then
            Separater1 = "."
        Else
            Separater1 = ","
        End If

        Dim s As New System.Text.RegularExpressions.Regex( _
        "\d+[" & Separater1 & "]\d+", _
        System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        'TextBox1.Text内で正規表現と一致する対象を1つ検索 
        Dim n As System.Text.RegularExpressions.Match = s.Match(Text1)

        GetDetail = n.Value

    End Function

    Function GetHistory(ByVal iString As String, Optional ByVal iUSAFlag As Boolean = False) As Collection

        GetHistory = New Collection

        Dim r As New System.Text.RegularExpressions.Regex("<.*?>")

        Dim String1 As String
        Dim Array1 As String()
        Dim Rows1 As String()

        If iUSAFlag = False Then
            Try
                String1 = iString.Substring(iString.IndexOf("年初来高値："))
            Catch
                String1 = iString.Substring(iString.IndexOf("為替時系列"))
            End Try
            String1 = Replace(String1, Environment.NewLine, "")

            Array1 = Split(String1, "table")

            Rows1 = Split(Array1(1), "</tr>")
        Else
            String1 = iString.Substring(iString.IndexOf("Adj Close*"))
            String1 = Replace(String1, Environment.NewLine, "")

            Array1 = Split(String1, "/table")

            Rows1 = Split(Array1(0), "</tr>")
        End If

        Dim Array2 As String()

        Dim Num1 As Long
        If iUSAFlag = False Then
            Num1 = UBound(Rows1) - 1
        Else
            Num1 = UBound(Rows1) - 2
        End If

        Dim i, j As Long
        For i = 1 To Num1
            Dim Text1 As String
            Text1 = ""
            Array2 = Split(Trim(Rows1(i)), "</td>")
            Text1 &= r.Replace(Array2(0), "")
            For j = 1 To UBound(Array2)
                Text1 &= ";" & r.Replace(Array2(j), "")
            Next
            GetHistory.Add(Text1)
        Next

    End Function
End Module
