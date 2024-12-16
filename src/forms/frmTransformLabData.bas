Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10380
    DatasheetFontHeight =11
    ItemSuffix =141
    Right =16500
    Bottom =11760
    RecSrcDt = Begin
        0xfe0110f4d4fbe540
    End
    Caption ="Transform lab data"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =9596
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2989
                    Top =8473
                    Width =733
                    Height =360
                    FontSize =12
                    TabIndex =2
                    Name ="txtMetadataEndNum"

                    LayoutCachedLeft =2989
                    LayoutCachedTop =8473
                    LayoutCachedWidth =3722
                    LayoutCachedHeight =8833
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =685
                    Top =8473
                    Width =2304
                    Height =314
                    ForeColor =0
                    Name ="Label15"
                    Caption ="Metadata end column #:"
                    LayoutCachedLeft =685
                    LayoutCachedTop =8473
                    LayoutCachedWidth =2989
                    LayoutCachedHeight =8787
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2996
                    Top =7988
                    Width =2880
                    Height =360
                    FontSize =12
                    TabIndex =1
                    Name ="txtRequestorName"

                    LayoutCachedLeft =2996
                    LayoutCachedTop =7988
                    LayoutCachedWidth =5876
                    LayoutCachedHeight =8348
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1379
                            Top =7989
                            Width =1623
                            Height =314
                            ForeColor =0
                            Name ="Label21"
                            Caption ="Requestor name:"
                            LayoutCachedLeft =1379
                            LayoutCachedTop =7989
                            LayoutCachedWidth =3002
                            LayoutCachedHeight =8303
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7505
                    Top =7080
                    Width =2520
                    Height =393
                    TabIndex =3
                    ForeColor =0
                    Name ="cmdTransformDataset"
                    Caption ="Transform dataset"
                    OnClick ="[Event Procedure]"
                    LeftPadding =66
                    RightPadding =79
                    BottomPadding =118

                    LayoutCachedLeft =7505
                    LayoutCachedTop =7080
                    LayoutCachedWidth =10025
                    LayoutCachedHeight =7473
                    ForeTint =100.0
                    BackThemeColorIndex =5
                    BackTint =100.0
                    BorderThemeColorIndex =5
                    BorderTint =100.0
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =24
                    QuickStyleMask =-1
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2999
                    Top =7506
                    Width =4320
                    Height =360
                    FontSize =12
                    Name ="txtLabName"
                    DefaultValue ="\"CCAL Water Analysis Laboratory\""

                    LayoutCachedLeft =2999
                    LayoutCachedTop =7506
                    LayoutCachedWidth =7319
                    LayoutCachedHeight =7866
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1979
                            Top =7506
                            Width =1008
                            Height =314
                            ForeColor =0
                            Name ="Label24"
                            Caption ="Lab name:"
                            LayoutCachedLeft =1979
                            LayoutCachedTop =7506
                            LayoutCachedWidth =2987
                            LayoutCachedHeight =7820
                            ForeTint =100.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =3827
                    Top =8473
                    Width =5826
                    Height =314
                    ForeColor =0
                    Name ="Label128"
                    Caption ="The last column, starting from the left, before the parameters."
                    LayoutCachedLeft =3827
                    LayoutCachedTop =8473
                    LayoutCachedWidth =9653
                    LayoutCachedHeight =8787
                    ForeTint =100.0
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2999
                    Top =7080
                    Width =4320
                    Height =327
                    FontSize =12
                    TabIndex =4
                    Name ="cboSourceTableName"
                    RowSourceType ="Value List"
                    RowSource ="source_2021_a;source_2021_b;source_table_a"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =2999
                    LayoutCachedTop =7080
                    LayoutCachedWidth =7319
                    LayoutCachedHeight =7407
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =120
                            Top =7081
                            Width =2893
                            Height =314
                            ForeColor =0
                            Name ="Label124"
                            Caption ="Lab data table name (SOURCE):"
                            LayoutCachedLeft =120
                            LayoutCachedTop =7081
                            LayoutCachedWidth =3013
                            LayoutCachedHeight =7395
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3001
                    Top =8969
                    TabIndex =5
                    Name ="chkSeparateSiteIDAndCode"
                    DefaultValue ="False"

                    LayoutCachedLeft =3001
                    LayoutCachedTop =8969
                    LayoutCachedWidth =3261
                    LayoutCachedHeight =9209
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3231
                            Top =8939
                            Width =5106
                            Height =537
                            ForeColor =0
                            Name ="Label132"
                            Caption ="Separate site ID and code from source column 'Site ID'.\015\012Requires source f"
                                "ormat like: 'DENA-001 Rock Creek'."
                            LayoutCachedLeft =3231
                            LayoutCachedTop =8939
                            LayoutCachedWidth =8337
                            LayoutCachedHeight =9476
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7566
                    Top =65
                    Width =2520
                    Height =1080
                    TabIndex =6
                    ForeColor =0
                    Name ="cmdShowMissingParameters"
                    Caption ="Show ANY source parameter(s) NOT in mapping table"
                    OnClick ="[Event Procedure]"
                    LeftPadding =65
                    RightPadding =79
                    BottomPadding =118

                    LayoutCachedLeft =7566
                    LayoutCachedTop =65
                    LayoutCachedWidth =10086
                    LayoutCachedHeight =1145
                    ForeTint =100.0
                    BackThemeColorIndex =6
                    BackTint =100.0
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    HoverThemeColorIndex =6
                    HoverTint =80.0
                    PressedThemeColorIndex =6
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =25
                    QuickStyleMask =-1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =87
                    Left =118
                    Top =118
                    Width =7384
                    Height =864
                    FontWeight =700
                    ForeColor =0
                    Name ="Label136"
                    Caption ="Step 1: \015\012If some source parameters are NOT in mapping table 'xref_map_col"
                        "umn_names', then add them to this table before transforming."
                    LayoutCachedLeft =118
                    LayoutCachedTop =118
                    LayoutCachedWidth =7502
                    LayoutCachedHeight =982
                    ForeTint =100.0
                End
                Begin Line
                    BorderWidth =4
                    OverlapFlags =85
                    Left =122
                    Top =3011
                    Width =10145
                    Name ="Line137"
                    LayoutCachedLeft =122
                    LayoutCachedTop =3011
                    LayoutCachedWidth =10267
                    LayoutCachedHeight =3011
                End
                Begin Label
                    OverlapFlags =85
                    Left =118
                    Top =3076
                    Width =10146
                    Height =3938
                    FontWeight =700
                    ForeColor =0
                    Name ="Label138"
                    Caption ="Step 2: \015\012- Select the SOURCE CCAL table (or table link) from drop-down.  "
                        "\015\012  - NOTE: The SOURCE recordset must have a field named 'RowID'. This fie"
                        "ld's values MUST be a unique\015\012     identifier for each recordset record in"
                        " 'rst'.\015\012- If necessary, enter the lab name and name of person that reques"
                        "ted the lab analysis.\015\012- Enter the LAST metadata column number (numbered l"
                        "eft to right, from the first column), from the source\015\012   table, into the "
                        "field 'Metadata end column #'. This column is before the first lab parameter col"
                        "umn.\015\012- Sometimes the site number and its name are combined in the source "
                        "'Site ID' column.\015\012  - If in this 'Site ID' is in this format (e.g. 'DENA-"
                        "001 Rock Creek'), checking 'Separate site ID and code will\015\012     cause thi"
                        "s utility to separate them into columns 'SiteID' and 'SiteCode' in the target ta"
                        "ble.\015\012- Click the 'Transform dataset' button to unpivot source table into "
                        "target table.\015\012  - NOTE: The target table is found in the list of tables i"
                        "n the navigation pane, and is formatted, for example,\015\012     as: 'stage_wat"
                        "er_chemistry_2024-02-08T10:18:58'. Each target table has a new timestamp in its "
                        "name, so \015\012     each target table is unique."
                    LayoutCachedLeft =118
                    LayoutCachedTop =3076
                    LayoutCachedWidth =10264
                    LayoutCachedHeight =7014
                    ForeTint =100.0
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =118
                    Top =1453
                    Width =10145
                    TabIndex =7
                    Name ="lstMissingParameters"
                    RowSourceType ="Value List"

                    LayoutCachedLeft =118
                    LayoutCachedTop =1453
                    LayoutCachedWidth =10263
                    LayoutCachedHeight =2893
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =118
                            Top =1087
                            Width =7530
                            Height =315
                            ForeColor =0
                            Name ="Label140"
                            Caption ="Missing parameter(s) in mapping table as compared with selected SOURCE table:"
                            LayoutCachedLeft =118
                            LayoutCachedTop =1087
                            LayoutCachedWidth =7648
                            LayoutCachedHeight =1402
                            ForeTint =100.0
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "frmTransformLabData.cls"
