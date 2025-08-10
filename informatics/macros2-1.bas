Attribute VB_Name = "Module1"
Sub Оформление_ячеек()
Attribute Оформление_ячеек.VB_ProcData.VB_Invoke_Func = "о\n14"
'
' Оформление_ячеек Макрос
'
' Сочетание клавиш: Ctrl+о
'
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Times New Roman"
        .FontStyle = "курсив"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .ColorIndex = 5
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
        .PatternTintAndShade = 0
    End With
    Application.Left = 103.75
    Application.Top = 97
End Sub
Sub копия_1лист()
Attribute копия_1лист.VB_ProcData.VB_Invoke_Func = "к\n14"
'
' копия_1лист Макрос
'
' Сочетание клавиш: Ctrl+к
'
    Sheets("Лист1").Select
    ActiveSheet.Buttons.Add(917.25, 9.75, 108, 75).Select
    Sheets("Лист1").Copy After:=Sheets(2)
End Sub
Sub фильтр_нумерация()
Attribute фильтр_нумерация.VB_ProcData.VB_Invoke_Func = "ф\n14"
'
' фильтр_нумерация Макрос
'
' Сочетание клавиш: Ctrl+ф
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("E1:E3").Select
    Selection.AutoFill Destination:=Range("E1:E6"), Type:=xlFillSeries
    Range("E1:E6").Select
    Range("E7:E11").Select
    Selection.AutoFill Destination:=Range("E1:E392"), Type:=xlFillDefault
    Range("E1:E392").Select
    ActiveWindow.SmallScroll Down:=-3
    ActiveWindow.ScrollRow = 352
    ActiveWindow.ScrollRow = 350
    ActiveWindow.ScrollRow = 347
    ActiveWindow.ScrollRow = 341
    ActiveWindow.ScrollRow = 333
    ActiveWindow.ScrollRow = 321
    ActiveWindow.ScrollRow = 315
    ActiveWindow.ScrollRow = 299
    ActiveWindow.ScrollRow = 282
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 244
    ActiveWindow.ScrollRow = 229
    ActiveWindow.ScrollRow = 204
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 169
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 1
    Columns("E:E").Select
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 104
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 116
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 118
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 133
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 143
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 149
    ActiveWindow.ScrollRow = 150
    ActiveWindow.ScrollRow = 151
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 153
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 158
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 161
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 166
    ActiveWindow.ScrollRow = 168
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 172
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 181
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 185
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 190
    ActiveWindow.ScrollRow = 192
    ActiveWindow.ScrollRow = 194
    ActiveWindow.ScrollRow = 196
    ActiveWindow.ScrollRow = 197
    ActiveWindow.ScrollRow = 198
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 201
    ActiveWindow.ScrollRow = 202
    ActiveWindow.ScrollRow = 203
    ActiveWindow.ScrollRow = 204
    ActiveWindow.ScrollRow = 206
    ActiveWindow.ScrollRow = 207
    ActiveWindow.ScrollRow = 208
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 210
    ActiveWindow.ScrollRow = 211
    ActiveWindow.ScrollRow = 212
    ActiveWindow.ScrollRow = 213
    ActiveWindow.ScrollRow = 214
    ActiveWindow.ScrollRow = 215
    ActiveWindow.ScrollRow = 216
    ActiveWindow.ScrollRow = 217
    ActiveWindow.ScrollRow = 218
    ActiveWindow.ScrollRow = 219
    ActiveWindow.ScrollRow = 220
    ActiveWindow.ScrollRow = 221
    ActiveWindow.ScrollRow = 222
    ActiveWindow.ScrollRow = 223
    ActiveWindow.ScrollRow = 224
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 226
    ActiveWindow.ScrollRow = 228
    ActiveWindow.ScrollRow = 229
    ActiveWindow.ScrollRow = 230
    ActiveWindow.ScrollRow = 231
    ActiveWindow.ScrollRow = 232
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 234
    ActiveWindow.ScrollRow = 235
    ActiveWindow.ScrollRow = 236
    ActiveWindow.ScrollRow = 237
    ActiveWindow.ScrollRow = 238
    ActiveWindow.ScrollRow = 239
    ActiveWindow.ScrollRow = 240
    ActiveWindow.ScrollRow = 241
    ActiveWindow.ScrollRow = 242
    ActiveWindow.ScrollRow = 243
    ActiveWindow.ScrollRow = 245
    ActiveWindow.ScrollRow = 246
    ActiveWindow.ScrollRow = 248
    ActiveWindow.ScrollRow = 249
    ActiveWindow.ScrollRow = 250
    ActiveWindow.ScrollRow = 251
    ActiveWindow.ScrollRow = 252
    ActiveWindow.ScrollRow = 253
    ActiveWindow.ScrollRow = 254
    ActiveWindow.ScrollRow = 255
    ActiveWindow.ScrollRow = 256
    ActiveWindow.ScrollRow = 257
    ActiveWindow.ScrollRow = 258
    ActiveWindow.ScrollRow = 259
    ActiveWindow.ScrollRow = 260
    ActiveWindow.ScrollRow = 261
    ActiveWindow.ScrollRow = 262
    ActiveWindow.ScrollRow = 263
    ActiveWindow.ScrollRow = 264
    ActiveWindow.ScrollRow = 265
    ActiveWindow.ScrollRow = 266
    ActiveWindow.ScrollRow = 267
    ActiveWindow.ScrollRow = 268
    ActiveWindow.ScrollRow = 270
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 272
    ActiveWindow.ScrollRow = 273
    ActiveWindow.ScrollRow = 274
    ActiveWindow.ScrollRow = 275
    ActiveWindow.ScrollRow = 276
    ActiveWindow.ScrollRow = 277
    ActiveWindow.ScrollRow = 278
    ActiveWindow.ScrollRow = 279
    ActiveWindow.ScrollRow = 280
    ActiveWindow.ScrollRow = 281
    ActiveWindow.ScrollRow = 282
    ActiveWindow.ScrollRow = 283
    ActiveWindow.ScrollRow = 285
    ActiveWindow.ScrollRow = 286
    ActiveWindow.ScrollRow = 287
    ActiveWindow.ScrollRow = 288
    ActiveWindow.ScrollRow = 289
    ActiveWindow.ScrollRow = 290
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 292
    ActiveWindow.ScrollRow = 293
    ActiveWindow.ScrollRow = 294
    ActiveWindow.ScrollRow = 295
    ActiveWindow.ScrollRow = 296
    ActiveWindow.ScrollRow = 297
    ActiveWindow.ScrollRow = 298
    ActiveWindow.ScrollRow = 299
    ActiveWindow.ScrollRow = 300
    ActiveWindow.ScrollRow = 301
    ActiveWindow.ScrollRow = 302
    ActiveWindow.ScrollRow = 303
    ActiveWindow.ScrollRow = 304
    ActiveWindow.ScrollRow = 306
    ActiveWindow.ScrollRow = 307
    ActiveWindow.ScrollRow = 308
    ActiveWindow.ScrollRow = 309
    ActiveWindow.ScrollRow = 310
    ActiveWindow.ScrollRow = 311
    ActiveWindow.ScrollRow = 312
    ActiveWindow.ScrollRow = 313
    ActiveWindow.ScrollRow = 314
    ActiveWindow.ScrollRow = 315
    ActiveWindow.ScrollRow = 316
    ActiveWindow.ScrollRow = 318
    ActiveWindow.ScrollRow = 320
    ActiveWindow.ScrollRow = 321
    ActiveWindow.ScrollRow = 323
    ActiveWindow.ScrollRow = 326
    ActiveWindow.ScrollRow = 329
    ActiveWindow.ScrollRow = 331
    ActiveWindow.ScrollRow = 332
    ActiveWindow.ScrollRow = 333
    ActiveWindow.ScrollRow = 334
    ActiveWindow.ScrollRow = 335
    ActiveWindow.ScrollRow = 336
    ActiveWindow.ScrollRow = 337
    ActiveWindow.ScrollRow = 339
    ActiveWindow.ScrollRow = 342
    ActiveWindow.ScrollRow = 343
    ActiveWindow.ScrollRow = 346
    ActiveWindow.ScrollRow = 348
    ActiveWindow.ScrollRow = 349
    ActiveWindow.ScrollRow = 350
    ActiveWindow.ScrollRow = 351
    ActiveWindow.ScrollRow = 352
    ActiveWindow.ScrollRow = 353
    ActiveWindow.ScrollRow = 354
    ActiveWindow.ScrollRow = 355
    ActiveWindow.SmallScroll Down:=21
    Range("E389:E392").Select
    Selection.AutoFill Destination:=Range("E389:E405"), Type:=xlFillDefault
    Range("E389:E405").Select
    Selection.AutoFill Destination:=Range("E389:E1234"), Type:=xlFillDefault
    Range("E389:E1234").Select
    ActiveWindow.ScrollRow = 1194
    ActiveWindow.ScrollRow = 1189
    ActiveWindow.ScrollRow = 1181
    ActiveWindow.ScrollRow = 1180
    ActiveWindow.ScrollRow = 1178
    ActiveWindow.ScrollRow = 1171
    ActiveWindow.ScrollRow = 1168
    ActiveWindow.ScrollRow = 1162
    ActiveWindow.ScrollRow = 1157
    ActiveWindow.ScrollRow = 1155
    ActiveWindow.ScrollRow = 1149
    ActiveWindow.ScrollRow = 1144
    ActiveWindow.ScrollRow = 1142
    ActiveWindow.ScrollRow = 1140
    ActiveWindow.ScrollRow = 1134
    ActiveWindow.ScrollRow = 1127
    ActiveWindow.ScrollRow = 1124
    ActiveWindow.ScrollRow = 1121
    ActiveWindow.ScrollRow = 1119
    ActiveWindow.ScrollRow = 1111
    ActiveWindow.ScrollRow = 1104
    ActiveWindow.ScrollRow = 1101
    ActiveWindow.ScrollRow = 1098
    ActiveWindow.ScrollRow = 1095
    ActiveWindow.ScrollRow = 1090
    ActiveWindow.ScrollRow = 1088
    ActiveWindow.ScrollRow = 1080
    ActiveWindow.ScrollRow = 1077
    ActiveWindow.ScrollRow = 1068
    ActiveWindow.ScrollRow = 1060
    ActiveWindow.ScrollRow = 1055
    ActiveWindow.ScrollRow = 1044
    ActiveWindow.ScrollRow = 1039
    ActiveWindow.ScrollRow = 1033
    ActiveWindow.ScrollRow = 1013
    ActiveWindow.ScrollRow = 1002
    ActiveWindow.ScrollRow = 980
    ActiveWindow.ScrollRow = 946
    ActiveWindow.ScrollRow = 866
    ActiveWindow.ScrollRow = 835
    ActiveWindow.ScrollRow = 758
    ActiveWindow.ScrollRow = 732
    ActiveWindow.ScrollRow = 693
    ActiveWindow.ScrollRow = 672
    ActiveWindow.ScrollRow = 652
    ActiveWindow.ScrollRow = 621
    ActiveWindow.ScrollRow = 603
    ActiveWindow.ScrollRow = 583
    ActiveWindow.ScrollRow = 536
    ActiveWindow.ScrollRow = 510
    ActiveWindow.ScrollRow = 486
    ActiveWindow.ScrollRow = 437
    ActiveWindow.ScrollRow = 276
    ActiveWindow.ScrollRow = 229
    ActiveWindow.ScrollRow = 205
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 161
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 1
    Columns("A:A").Select
    Application.Run "Книга3!фильтр_нумерация"
    Columns("A:A").Select
    Application.Run "Книга3!фильтр_нумерация"
    Columns("A:A").Select
    Selection.Copy
    Range("F1").Select
    ActiveSheet.Paste
    Selection.ColumnWidth = 24.14
    Range("H5").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Columns("F:F").Select
    ActiveWorkbook.Worksheets("Лист10").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Лист10").Sort.SortFields.Add2 Key:=Range("F1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Лист10").Sort
        .SetRange Range("F1:F1234")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
