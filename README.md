# jpbz
private files
'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2014/04/03 09:40:39
'-- Author: 
'-- Comment:
'ret = FileNameGet("ANY", "FileRead", ScriptReadPath, "KS2 (*.ks2),*.ks2"," ＫＳ２形式データファイルを選択してくだいさい。")

'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

    '※今のところ未使用
    '■変更する変数を先頭に持ってくる■
    Dim  Const_ChannelName    'サンプリングするチャンネルの名簿
    Dim  Const_SampleHz       'サンプリングする周波数、名簿に合わせて周波数をセットする必要がある
    Dim  Const_TakeTime       'サンプリングする時間、サンプリングの終わり時間を指定(0～αまで)
    Dim  Const_MaxLength      'サンプリングする最大の長さ
    Dim  Const_WaveXForm      'WaveXFormきざみ
    Dim Const_CopyChName      'コピーするチャンネルの名前
    ''サンプリングするチャンネルの名簿の値をセットする
    ' 2014/12/26 桧山用設定に変更　①～⑤主軸　⑥高速軸　⑦中速軸　⑭回転数
    Const_ChannelName = array( _
    "圧電1","圧電2","圧電3","圧電4","圧電5","圧電6","圧電7","圧電8","圧電9", "ｻｰﾎﾞﾅｾﾙNS","ｻｰﾎﾞﾅｾﾙEW","ｻｰﾎﾞﾀﾜｰNS","ｻｰﾎﾞﾀﾜｰEW","回転計")

    'サンプリングする周波数の値のセット
    '                        ①    ②    ③　　④    ⑤　  ⑥　  ⑦    ⑧   ⑨    NS    EW    NS    EW   RPM　　
    Const_SampleHz = array( 0100, 0100, 0100, 0100, 0100, 4000, 1000, 400, 4000, 0050, 0050, 0050, 0050, 0100 )
    
    'サンプリングする時間の値のセット
    '                         ①     ②    ③      ④     ⑤     ⑥      ⑦    ⑧     ⑨     NS    EW    NS    EW    RPM　　
    Const_TakeTime = array( 032.0, 032.0, 032.0, 032.0, 032.0,  001.6, 006.4, 0016, 001.6,  64.0, 64.0, 64.0, 64.0,  64.0 )
    '                         ①    ②    ③　　④    ⑤　  ⑥　  ⑦    ⑧   ⑨    NS    EW    NS    EW   RPM　
    'Const_MaxLength = array( 5200, 5200, 5200, 5200, 5200, 3200, 3200, 3200, 3200, 3200, 3200, 3200, 3200, 3200)
    Const_MaxLength = array( 1600, 1600, 1600, 1600, 1600, 3200, 3200, 3200, 3200, 3200, 3200, 3200, 3200, 3200)

'                                ①       ②      ③　　   ④       ⑤　    ⑥　   ⑦     ⑧      ⑨     NS     EW     NS     EW    RPM　
    Const_WaveXForm  = array( 0.03125, 0.03125, 0.03125, 0.03125, 0.03125, 0.625, 0.15625, 0.625, 0.625, 0.625, 0.625, 0.625, 0.625, 0.015625 )

    Const_CopyChName = "圧電6" 

    '■変更する変数を先頭に持ってくる■


    Dim OriginalChGrp     'リサンプリング元データ
    Dim ResampleChGrp     'リサンプリング先データ

    Dim FileList          'ダイアログで指定したフォルダの中にあるファイルのリスト
    Dim Name              'リサンプリング先ファイルの名前
    Dim Ret               'ダイアログで選択されたコマンドの返り値
    Dim n                 'カウント変数
    Dim OutPath           'リサンプリングデータを作成するフォルダのパス

    Dim Edit              '編集したかどうかのフラグ?
    dim Ary_date()

    '//  set edit log
    Edit = False

    Ret = pathDlgShow("KS2ファイルが格納されたフォルダを選択してください","C:\")
    if Ret = "IDOk" then
        OutPath = outputPath
        '// obtain list of file
        FileList = DirListGet( OutPath, "*.ks2" , "FileName")
        redim preserve ary_date(Ubound( FileList ) )
        If IsArray( FileList ) then
            for n= Lbound( FileList ) To Ubound( FileList ) 
                '// read data　リサンプリングするデータを読み込む
                Name = FileList( n )
                Call Data.Root.Clear()
                Call DataFileLoad( OutPath & Name & ".ks2" , "Kyowa_KS2" ,"Load")
                set OriginalChGrp = Data.Root.ActiveChannelGroup
                ary_date(n) = originalchgrp.Properties.Item("Meas_start_date").Value
                '// create resampe group　リサンプリングデータのグループを作成する(画面右側に表示されるグループ)
                Call CreateResample()
                    
                '// shadow by 圧電6　指定したチャネルを3分割する
                'Call CreateShadow()    '2014/12/16打ち合わせの際に一つのセンサーに付き1計測と話があったので、コピーは作成しない    
                    
                '// resample preparation　各チャンネルの周波数、時間、ローパスを指定してリサンプリングする
                Call  PreparationResample()
                        
                '// resample parameter edit
                if Edit = False then
                        
                    'Call SUDDlgShow( "Resample", scriptReadPath & "Dialog.sud")
                    Edit = True
                end if
                        
                '// filter process　フィルタリングする
                Call FilterProcess()

                '// resample process
                Call ResampleProcess()
                        
                'リサンプリングしたデータをTDMに保存する
                FolderCreate(OutPath & "TDM")   'TDMフォルダを作成する
                Call DataFileSaveSel( OutPath & "TDM\" & Name & ".TDM","TDM", ResampleChGrp )
            
                '// delete original group　リサンプリング元のデータを右の画面から削除する
                Data.Root.ChannelGroups.Remove( OriginalChGrp.Name )  
                Call Data.Root.Clear
            Next       
        End if    
        'リサンプリングが終わったら、TDMファイルを読み込み、ACC、FFT、ENVFFTのファイルを作成する
        FileList = DirListGet( OutPath & "TDM\" , "*.TDM" , "FileName")
        If IsArray( FileList ) then

            For n= Lbound( FileList ) To Ubound( FileList ) 
                Name = FileList( n )
                Call DataFileLoad( OutPath & "TDM\" & Name, "TDM" ,"Load")
            Next
        End If
        For n = 1 to data.Root.ChannelGroups(1).Channels.Count - 1 
            Call MakeACC(n)          'ACC.csvを作成する
            Call MakeFFT(n)          'FFT.csvを作成する
            Call MakeEnvFFT(n)       'EnvFFT.csvを作成する
        Next
        Call data.Root.Clear     '読み込んだデータを削除する
        MsgBox("処理完了")       '処理完了のメッセージを出す
    End If


'------------------------------------------------------------------------------'
'関数名:createResample                                                         '
'機能　:リサンプリングデータのグループを作成する(画面右側に表示されるグループ) '
'引数　:                                                                       '
'返り値:                                                                       '
'------------------------------------------------------------------------------'
Sub CreateResample()

    Dim GrpName
    GrpName = Data.Root.ActiveChannelGroup.Name
    Set ResampleChGrp =  Data.Root.ChannelGroups.Add( GrpName & "_resample")
    Call MsgLineDisp( GrpName )

End Sub

'------------------------------------------------------------------------------'
'関数名:createShadow                                                           '
'機能　:指定したチャネルを3分割する                                            '
'引数　:OriginalGrp「コピー元のグループ」　                                    '
'返り値:                                                                       '
'------------------------------------------------------------------------------'
Sub CreateShadow()
    Dim CONST_EXT
    Dim GrpFrom
    Dim GrpTo
    Dim n
    Dim m
    CONST_EXT = Array( "a", "b", "c" )

    For n= 1 To OriginalChGrp.Channels.Count
            
        Set GrpFrom = OriginalChGrp.Channels( n )
        '★チャンネルのコピーを作成したいチャンネル名を記述する(常数の定義はスクリプトの先頭にあるので、値はそちらを変更すること)
        If OriginalChGrp.Channels( n ).Name = Const_CopyChName Then
            
        '// create shadow
            For m=1 To 3
                Set GrpTo = OriginalChGrp.Channels.Add( GrpFrom.Name, DataTypeChnFloat64)                 
                Call ChnCopy(  GrpFrom, GrpTo )       'チャンネルのコピー
                Call ChnPropCopy( GrpFrom, GrpTo )    'プロパティのコピー
                GrpTo.Name = GrpFrom.Name & CONST_EXT( m-1 )  'コピー先のチャンネルの名前変更
            Next 
                    
            '// move up
            For m=1 To 3
                    
                Set GrpTo = OriginalChGrp.Channels( GrpFrom.Name & CONST_EXT( m-1 ) ) 'コピー先のチャンネルをセット                 
                
                Call Data.Move( GrpTo, OriginalChGrp.Channels, n + m )  'コピーしたチャンネルを所定の位置に移動する
            Next                                                          '圧電6aなら圧電6の下に移動
                    
            '// delete original
            Call ChnDelete( GrpFrom ) 'コピー元のチャンネルを削除                         
        End If
    Next
End Sub

'------------------------------------------------------------------------------'
'関数名:preparationResample                                                    '
'機能　:各チャンネルの周波数、時間、ローパスを指定してリサンプリングする       '
'引数　:                                 　                                    '
'返り値:                                                                       '
'------------------------------------------------------------------------------'
Sub  PreparationResample()
    Dim ChannelFrom           'サンプリングするチャンネルのオブジェクト
    Dim ChannelName           'リサンプリングするチャンネルの名前
    Dim SampleHz              'セットする周波数(const_channelNameの値を入れる器の役割)
    Dim TakeTime              'セットする時間(const_takeTimeの値を入れる器の役割)

    Dim n                     'カウント変数
    Dim m                     'カウント変数
    
    'リサンプリングするチャンネルの数だけループする
    For n= 1 To OriginalChGrp.Channels.Count     
    
        Set ChannelFrom = OriginalChGrp.Channels( n )   'チャンネルのセット
            
        '// csan name 　チャンネルの名簿の数だけループする　
        For m = 0 to Ubound( Const_ChannelName ) 
          
            ChannelName = ChannelFrom.Properties("Name").Value   'リサンプリングするチャンネルの名前を取得
            If ChannelName = Const_ChannelName( m ) Then     'リサンプリングするチャンネルの名前とチャンネル名簿の名前が一致した場合
                    
                SampleHz = Const_SampleHz( m )    'リサンプリングする周波数をセット
                TakeTime  = Const_TakeTime( m )   'リサンプリングする時間をセット
                            
                'リサンプリングするチャンネルのプロパティにリサンプリング条件をセットする
                '// resampleHZ(周波数)
                Call ChannelFrom.Properties.Add("ResampleHZ",SampleHz,DataTypeInt32) 
            
                '// takeTimeSEC(時間)
                'takeTime = ChannelFrom.Properties("wf_increment").Value * ChannelFrom.Properties("lengthmax").Value
                Call ChannelFrom.Properties.Add("TakeTimeSEC", TakeTime,DataTypeFloat32)                            
 
                '// lowpassHz(ローパス)
                Call ChannelFrom.Properties.Add("LowPassHz", SampleHz/2,DataTypeInt32) 
                            
                Exit For
            End If
        Next                
    Next
End Sub

'------------------------------------------------------------------------------'
'関数名:filterProcess                                                          '
'機能　:フィルタリングする                                                     '
'引数　:                                 　                                    '
'返り値:                                                                       '
'------------------------------------------------------------------------------'
Sub  FilterProcess()
    Dim n             'カウント変数
    Dim ChannelFrom   'リサンプリングするチャンネル
    'リサンプリングするチャンネルの数だけループする
    For n= 1 To OriginalChGrp.Channels.Count

        '// obtain  lowpass frequency
        Set ChannelFrom = OriginalChGrp.Channels( n )   'チャンネルのセット             
        'チャンネル名が回転計でない場合
        If ChannelFrom.Properties("name").Value <> "回転計"  Then  
            'ローパスが0でない場合
            If ChannelFrom.Properties("lowpassHz").Value <> 0 Then
                '各種プロパティをセットする
                FiltStruc = "IIR"
                FiltStyle = "Bessel"
                FiltType = "Low Pass"
                FiltDegree = 2
                FiltLimit = ChannelFrom.Properties("LowPassHz").Value          
                FiltLowLimit = 50
                FiltUppLimit = 0
                FiltWave = 1.2
                FiltSamples = 25         
                FiltWndFct = "Hamming"
                FiltZeroPhase = 0
                FiltCorrection = 0
                Call ChnFiltCalc("",ChannelFrom, ChannelFrom, _
                                 FiltStruc, FiltStyle, FiltType, FiltDegree,Filtlimit,_
                                 FiltLowLimit, FiltUppLimit, FiltWave, FiltSamples, _
                                 FiltWndFct, FiltZeroPhase, FiltCorrection )
            End If                                            
        End If
    Next
End Sub

'------------------------------------------------------------------------------'
'関数名:resampleProcess                                                        '
'機能　:                                                                       '
'引数　:                                 　                                    '
'返り値:                                                                       '
'------------------------------------------------------------------------------'
sub  ResampleProcess()
    Dim ChannelFrom
    Dim ChannelTo

    Dim Ret
    Dim n
    Dim m

    Dim MicroSec
    Dim Size

    Call ResampleChGrp.Activate()   'リサンプリングチャンネルをアクティブにする
    'リサンプリングするチャンネルの数だけループする
    For n= 1 To OriginalChGrp.Channels.Count
        
        '// obtain  resample frequency
        Set ChannelFrom = OriginalChGrp.Channels( n )   'リサンプリングするチャンネルのセット             
        'リサンプリング周波数が0でない場合
        If ChannelFrom.Properties("resampleHZ").Value <> 0 then
                
            '// convert  from frequency to microsec
            MicroSec = Cint( ( 1/ ChannelFrom.Properties("resampleHZ").Value ) * 1000000 )
            Size =   ChannelFrom.Properties("lengthmax").Value
                        
            '// create time channel
            Size = Size * ( 1 / ( MicroSec / 100 ) )        '//  recording is 100 micro sec
            Call ChnGenTime("/TimeGenerated","MicroSecond",0,0,MicroSec,"StartStepNo",Size)
                
            '// resample
            Call ChnMapLinCalc("",ChannelFrom,"/TimeGenerated","/LinearMapped",1,"Const. Value",NOVALUE,"Analogue")
            Call ChnToWfChn("/TimeGenerated","/LinearMapped",1,"WfXRelative")
                        
            '// set name & properties
            Set ChannelTo = ResampleChGrp.Channels.Item("LinearMapped")
            ChannelTo.Name = ChannelFrom.Name
            Call ChnPropCopy( ChannelFrom, ChannelTo )                         
        Else
            '// through
            Set ChannelTo = ResampleChGrp.Channels.Add( ChannelFrom.Name, DataTypeChnFloat64)                 
            Call ChnCopy(ChannelFrom, ChannelTo )                    
        End If
                
        '// cut
        If ChannelTo.Properties("TakeTimeSEC").Value <> 0 Then
                
            ChnRow = Clng( CDbl(ChannelTo.Properties("takeTimeSEC").Value) / ChannelTo.Properties("Wf_Increment").Value )
            If ChannelTo.Properties("lengthmax").Value - ChnRow > 0 Then        
                ValNo = ChannelTo.Properties("lengthmax").Value - ChnRow                  
                Call DataBlDel( ChannelTo, ChnRow, ValNo)
            End If
        End If   
    Next
End Sub

'------------------------------------------------------------------------------'
'関数名:MakeACC                                                                '
'機能　:ACCファイルを作成する                                                  '
'引数　:                                 　                                    '
'返り値:                                                                       '
'------------------------------------------------------------------------------'
sub MakeACC(ByVal ChNo)
    Dim LineNo
    Dim IntMyHandle       '書き込み用のファイル
    Dim IntMyText
    Dim IntMyError
    Dim SamplNo
    Dim Buffer
    Dim i
    'チャンネル単位でデータをテキスト書き出しし、データ配列は以下のとおりとする。
    'time, ｸﾞﾙｰﾌﾟ1, ｸﾞﾙｰﾌﾟ2, ｸﾞﾙｰﾌﾟ3, ｸﾞﾙｰﾌﾟ4, DataPortalで表示されている上から順に横書きとする。DataPortalはHDD内でのファイル順
    '　0.001, 
    '　0.002
    '　0.003
    '　・・・

    '書き込み＆作成用テキストファイルのオープン  
    FolderCreate(outpath & "ACC" )    'フォルダを作成する(既に同名のフォルダがある場合は何もしない)
    IntMyHandle = TextFileOpen(OutPath & "ACC" & "\acc(" & ChNo & ").csv",TFCreate OR TFWrite OR TFANSI) 
                           
    '1行目の項目表示。この項目表示でグループ名が記録されるのでテキスト後のデータの識別ができる。
    Buffer = "Time" '1行1列目の文字

    '上記のTimeに続いて、1行の文字列を作る
    For i = 1 To Data.Root.ChannelGroups.Count      '表示チャンネル数が入る（Q要確認：DataPortalの個数が入るのか？）
        Buffer = Buffer & "," & Data.Root.ChannelGroups(i).Name '
        'Buffer = Buffer & "," & ary_date(i - 1)
    Next
    IntMyText= TextFileWriteLn(IntMyHandle, Buffer)
    Buffer = ""
    For i = 1 To Data.Root.ChannelGroups.Count      '表示チャンネル数が入る（Q要確認：DataPortalの個数が入るのか？）
        'Buffer = Buffer & "," & Data.Root.ChannelGroups(i).Name '
        Buffer = Buffer & "," & ary_date(i - 1)
    Next
    '↑→Q：普通にWrite、","で書き出されないのか。何か途中テキストファイル書き出しできないか。
    'TextfileWriteLn(intMyHandle,"," & Data.Root.ChannelGroups(i).Name )とかの記述で

    '上記の1行目用文字列を一気に書き出す
    IntMyText= TextFileWriteLn(IntMyHandle, Buffer)

    For SamplNo=1 To Const_MaxLength(ChNo - 1)   '★数はプロパティ（長さ）を確認して入力
        Buffer = (SamplNo-1) * Const_WaveXForm(ChNo - 1)   '★きざみはWaveform X間隔で確認して入力。"-1"は初期値がゼロのため。
    
        For i=1 To Data.Root.ChannelGroups.Count       'グループ数を入力 
                  '↑→Q："Data.Root.ChannelGroups.Count"ではだめなのか？？
      
            Buffer = Buffer & "," & Data.Root.ChannelGroups(i).Channels(ChNo).Values(SamplNo)
                    '★()内はチャンネル番号.チャンネルごとに生成する。
                    'Q:チャンネル番号はabcも含めた通し番号
                    'グループ～チャンネル～Ｘ値の順で読み出されるようだ。
                    'ｸﾞﾙｰﾌﾟはPortal表示順、チャンネルも表示順（圧電①、、）
                    'Q:途中からチャンネルが増減した場合は？    
        Next
        '上記同様、１行分を書き出していく。
        IntMyText= TextFileWriteLn(IntMyHandle, Buffer)
    Next 
    '書き出したテキストファイルをクローズ。
    IntMyError = TextFileClose(IntMyHandle)
    'Q:横軸の最後がきっちりで終わらないのだが。
End Sub

'------------------------------------------------------------------------------'
'関数名:MakeFFT                                                                '
'機能　:FFTファイルを作成する                                                  '
'引数　:                                 　                                    '
'返り値:                                                                       '
'------------------------------------------------------------------------------'
Sub MakeFFT(ByVal ChNo)
    'FFTなどの結果はPortalの最後部付加のため、グループで一纏め生成。新グループ作成・FFT群生成用
    Call Data.Root.ChannelGroups.Add("FFT").Activate()  

    '1チャンネルFFTの諸条件値を自動マクロで生成し、予め以下のように貼り付けておく

    FFTIndexChn      = 0
    FFTIntervUser    = "NumberStartOverl"
    FFTIntervPara(1) = 1
    FFTIntervPara(2) = 6400
    FFTIntervPara(3) = 1
    FFTIntervOverl   = 0
    FFTNoV           = 0
    FFTWndFct        = "Hanning"
    FFTWndPara       = 10
    'FFTWndChn        = "[1]/圧電1"
    FFTWndChn        = "[" & ChNo & "]/" &  Const_ChannelName(ChNo - 1) 
    FFTWndCorrectTyp = "Periodic"
    FFTAverageType   = "No"
    FFTAmplFirst     = "Amplitude"
    FFTAmpl          = 1
    FFTAmplType      = "Ampl.Peak"
    FFTCalc          = 0
    FFTAmplExt       = "No"
    FFTPhase         = 0
    FFTCepstrum      = 0

    '計算するチャンネルにしたがって、FFTWndChnはける必要があるのか？。他にも変更する箇所は？
    'Q： "[1]/圧電1"は自動マクロ時のダミーと思っておけばよいとのこと。

    Dim LineNo
    Dim IntMyHandle
    Dim IntMyText
    Dim IntMyError
    Dim SamplNo
    Dim Buffer
    Dim i

    'グループ（計測）数（毎に）FFTを全部生成する。
    For i = 1 To Data.Root.ChannelGroups.Count-1  
          '↑前出で生成したFFTグループは無視するため"-1"とする。
        Call ChnFFT1("",Data.Root.ChannelGroups(i).Channels(ChNo))   '★Ch番号を入力
        '↑「ChnFFT1」は上記の各変数を読み込んだ条件で算出する
        '↑チャンネル番号を変えてランさせる。チャンネルごとに全グループ（計測）をFFT生成。
        '↑このとき、生成されたFFTは、FFTグループ内に登録される。.Activateが効いているため？(Q)
        '↑Portalでのチャンネルの順番に生成（Q)。
    Next

    '書き出しファイルをの作成、open
    FolderCreate(outpath & "FFT" )    'フォルダを作成する(既に同名のフォルダがある場合は何もしない)
    IntMyHandle = TextFileOpen(OutPath & "FFT" & "\FFT(" & ChNo & ").csv",TfCreate OR TfWrite OR TfANSI) 
    Buffer = "Freq"   '1行1列目
    For i = 1 To Data.Root.ChannelGroups.Count-1          '上記同様
        'Buffer = Buffer & "," & Data.Root.ChannelGroups(i).Name   '1行目のグループ(計測番号)の書き出し。
        Buffer = Buffer & "," & ary_date(i - 1)
    Next
    IntMyText= TextFileWriteLn(IntMyHandle, Buffer) '1行目書き込み完了

    'FFT数値データの書き出し
    For SamplNo=1 to Const_MaxLength(ChNo - 1) '数はプロパティを確認して入力
                  '↑↓★手動でAmplitudeを一個作って各チャンネルごとに確認要
        Buffer = (SamplNo-1) * Const_WaveXForm(ChNo - 1)      '★waveform X間隔を参照   
    
        'グループ数を入力 ここでは生成したFFT数=計測数=現Portal表示のグループ数-1
        For i=1 to Data.Root.ChannelGroups.Count-1       
      
            Buffer = Buffer & "," & Data.Root.ChannelGroups("FFT").Channels(i).Values(SamplNo)
            ' "FFT"グループの、FFT各生成ファイル(チャンネル）の、X値上のY値を追加する。
        Next
    
        IntMyText= TextFileWriteLn(IntMyHandle, Buffer)
    Next 
  
    IntMyError = TextFileClose(IntMyHandle)
    Call Data.Root.ChannelGroups.Remove("FFT")  
End Sub

'------------------------------------------------------------------------------'
'関数名:MakeEnvFFT                                                             '
'機能　:EnvFFTファイルを作成する                                               '
'引数　:                                 　                                    '
'返り値:                                                                       '
'------------------------------------------------------------------------------'
sub MakeEnvFFT(ByVal ChNo)
    Dim LineNo
    Dim IntMyHandle
    Dim IntMyText
    Dim IntMyError
    Dim SamplNo
    Dim Buffer
    Dim i

    'FFTなどの結果はPortalの最後部付加のため、グループで一纏め生成。新グループ作成・FFT群生成用

    Call Data.Root.ChannelGroups.Add("FFT").Activate()  

    '全グループで指定チャンネルのenvFFTを作り、ｸﾞﾙｰﾌﾟ「FFT」に蓄積する。
    For i = 1 To Data.Root.ChannelGroups.Count-1

        '加速度波形ひっくり返し。計算機機能より。履歴で見れる。
        'Call Calculate("Ch("""")=IIF(Ch(""[" & CStr(i) & "]/圧電1"")<0,-1*Ch(""[" & CStr(i) & "]/圧電1""),Ch(""[" & CStr(i) & "]/圧電1""))",NULL,NULL,"")
        Call Calculate("Ch("""")=IIF(Ch(""[" & CStr(i) & "]/" & Const_ChannelName(ChNo - 1) & """)<0,-1*Ch(""[" & CStr(i) & "]/" & Const_ChannelName(ChNo - 1) & """),Ch(""[" & CStr(i) & "]/" & Const_ChannelName(ChNo - 1) & """))",NULL,NULL,"")

        '★チャンネル番号での設定でなく自動マクロコードで設定された変数「圧電～」で指定。

        'エンベロープ　フィッティング機能を利用
        Call ChnEnvelopes("","FFT/Calculated","/UpperEnvelopeX", _
                          "/UpperEnvelopeY","/LowerEnvelopeX","/LowerEnvelopeY",0.2083)
                                                          '★上記の値で包絡度を確定↑。
                                                          '振って感度分析すべし。
        '基本演算ーチャンネル関数、Timegenerated生成                    
        Call ChnFromWfXGen("FFT/Calculated","/TimeGenerated","WfXRelative")
                  
        'フィッティングーリニアマッピング
        Call ChnMapLinCalc("FFT/UpperEnvelopeX","FFT/UpperEnvelopeY","FFT/TimeGenerated","/LinearMapped",1,"const. value",NOVALUE,"analogue")
        'ＦＦＴ

        FFTIndexChn      = 0
        FFTIntervUser    = "NumberStartOverl"
        FFTIntervPara(1) = 1
        FFTIntervPara(2) = 6400
        FFTIntervPara(3) = 1
        FFTIntervOverl   = 0
        FFTNoV           = 0
        FFTWndFct        = "Hanning"
        FFTWndPara       = 10
'        FFTWndChn        = "[1]/" & Const_ChannelName(ChNo - 1)
        FFTWndChn        = "[" & ChNo & "]/" &  Const_ChannelName(ChNo - 1) 
        FFTWndCorrectTyp = "No"
        FFTAverageType   = "No"
        FFTAmplFirst     = "Amplitude"
        FFTAmpl          = 1
        FFTAmplType      = "Ampl.Peak"
        FFTCalc          = 0
        FFTAmplExt       = "No"
        FFTPhase         = 0
        FFTCepstrum      = 0
        Call ChnFFT1("FFT/TimeGenerated","FFT/LinearMapped")
        Call Data.Root.ChannelGroups("FFT").Channels.Remove("Calculated")
        Call Data.Root.ChannelGroups("FFT").Channels.Remove("UpperEnvelopeX")
        Call Data.Root.ChannelGroups("FFT").Channels.Remove("UpperEnvelopeY")
        Call Data.Root.ChannelGroups("FFT").Channels.Remove("LowerEnvelopeX")
        Call Data.Root.ChannelGroups("FFT").Channels.Remove("LowerEnvelopeY")
        Call Data.Root.ChannelGroups("FFT").Channels.Remove("TimeGenerated")
        Call Data.Root.ChannelGroups("FFT").Channels.Remove("LinearMapped")
        If i>1 Then
            Call Data.Root.ChannelGroups("FFT").Channels.Remove("Frequency1")
        End If
    Next

    '書き出し
    FolderCreate(outpath & "ENVFFT" )    'フォルダを作成する(既に同名のフォルダがある場合は何もしない)
    IntMyHandle = TextFileOpen(OutPath & "ENVFFT" & "\EnvFFT(" & ChNo &  ").csv",TfCreate OR TfWrite OR TfANSI)       
    Buffer = "Freq"   '1行1列目

    For i = 1 To Data.Root.ChannelGroups.Count-1          '上記同様
        'Buffer = Buffer & "," & Data.Root.ChannelGroups(i).Name   '1行目のグループ(計測番号)名の書き出し。
        Buffer = Buffer & "," & ary_date(i - 1)
    Next

    IntMyText= TextFileWriteLn(IntMyHandle, Buffer) '1行目書き込み完了

    'FFT数値データの書き出し
    For SamplNo=1 To Const_MaxLength(ChNo - 1) '★数はプロパティを確認して入力 
        'FFTのＸ軸の数値がチャンネル１に入っているため、これを書き出してから、各チャンネルのＹ値をループする。
        Buffer=Data.Root.ChannelGroups("FFT").Channels(1).Values(SamplNo)
        For i=2 To Data.Root.ChannelGroups.Count      
            Buffer = Buffer & "," & Data.Root.ChannelGroups("FFT").Channels(i).Values(SamplNo)
            '"FFT"グループの、FFT各生成ファイル(チャンネル）の、X値上のY値を追加する。
        Next
        IntMyText= TextFileWriteLn(IntMyHandle, Buffer)
    Next 
    IntMyError = TextFileClose(IntMyHandle)
    Call Data.Root.ChannelGroups.Remove("FFT")  
End Sub
