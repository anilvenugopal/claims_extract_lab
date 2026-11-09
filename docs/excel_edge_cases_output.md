# Workbook: excel_edge_cases.xlsx

> **How to read this file.** Each sheet is split into blocks. In every table, the column letters are across the top and the row numbers are down the left side, exactly as in Excel. To cite a value, give the sheet name and the cell, for example `'Q3 Report'!B7`. Empty cells are blank. Rows and columns that are completely empty are left out, so row numbers can skip. Merged cells show their value only in the first cell of the range (listed under the table). Cell comments are listed under the table as `Comment on B2 (by author): text`. Numbers in date-like columns that Excel stores as plain numbers are shown as `2026-10-05 (serial 46300)`: the date, then the original number. Hidden sheets, rows and columns are marked.

Sheets in order: Summary (visible), 'Q3 Report' (visible), 'Multi Tables' (visible), 'No Header Data' (visible), 'Hidden Rates' (hidden), 'Config Secret' (veryHidden), 'Hidden Rows Cols' (visible), Empty (visible), 'Big Log' (visible), Messy (visible), 'Merged Cells' (visible), Comments (visible), Dates (visible), 'General Dates' (visible)

## Sheet Summary — visible

### Table 1 — Summary!A1:C5

| | A | B | C |
|---|---|---|---|
| 1 | Product | Units | Price |
| 2 | Widget | 120 | 9.99 |
| 3 | Gadget | 75 | 24.5 |
| 4 | Gizmo | 30 | 0.3 |
| 5 | Doohickey | 200 | 4 |

## Sheet 'Q3 Report' — visible

Regions on this sheet: table A1:E3, table A5:E11, table A13:E14

### Table 1 — 'Q3 Report'!A1:E3

| | A | B |
|---|---|---|
| 1 | Q3 Regional Sales Report |  |
| 2 | Prepared by: Finance Team |  |
| 3 | Date: | 2026-10-01 |

- Merged cells (value sits in the first cell of the range): A1:E1 ("Q3 Regional Sales Report")

### Table 2 — 'Q3 Report'!A5:E11

| | A | B | C | D | E |
|---|---|---|---|---|---|
| 5 | Region | Q3 Revenue ($) |  |  | Total |
| 6 |  | Jul | Aug | Sep |  |
| 7 | East | 1200 | 1350 | 1100 | 3650 |
| 8 | West | 900 | 950 | 1010 | 2860 |
| 9 | North | 400 | 420 | 380 | 1200 |
| 10 | South | 760 | 800 | 845 | 2405 |
| 11 | Total | 3260 | 3520 | 3335 | 10115 |

- Merged cells (value sits in the first cell of the range): A5:A6 ("Region"); B5:D5 ("Q3 Revenue ($)"); E5:E6 ("Total")

### Table 3 — 'Q3 Report'!A13:E14

| | A |
|---|---|
| 13 | Note: figures are unaudited \| source: ERP export.<br>Second line of the note. |
| 14 | Footnote spanning the table width, merged across A to E. |

- Merged cells (value sits in the first cell of the range): A14:E14 ("Footnote spanning the table width, merg…")

## Sheet 'Multi Tables' — visible

Regions on this sheet: table B2:D6, table H3:I5, table B11:D14, cell K20, table B25:C26

### Table 1 — 'Multi Tables'!B2:D6

| | B | C | D |
|---|---|---|---|
| 2 | Headcount by department |  |  |
| 3 | Dept | Heads | Open roles |
| 4 | Eng | 40 | 5 |
| 5 | Sales | 25 | 2 |
| 6 | Ops | 12 | 0 |

### Table 2 — 'Multi Tables'!H3:I5

| | H | I |
|---|---|---|
| 3 | Dept | Budget |
| 4 | Eng | 4000000 |
| 5 | Sales | 1500000 |

### Table 3 — 'Multi Tables'!B11:D14

| | B | C | D |
|---|---|---|---|
| 11 | 1 | 2 | 3 |
| 12 | 4 | 5 | 6 |
| 13 | 7 | 8 | 9 |
| 14 | 10 | 11 | 12 |

**K20** (single cell): See appendix B

### Table 4 — 'Multi Tables'!B25:C26

| | B | C |
|---|---|---|
| 25 | Owner | J. Rivera |
| 26 | Reviewed | 2026-09-30 14:30:00 |

## Sheet 'No Header Data' — visible

### Table 1 — 'No Header Data'!C3:F6

| | C | D | E | F |
|---|---|---|---|---|
| 3 | 1 | 2.5 | 3 | TRUE |
| 4 |  | 4 | 5.25 | FALSE |
| 5 | 7 |  | 2026-01-15 | n/a |
| 6 | 10 | 11 | 12 | 13 |

## Sheet 'Hidden Rates' — HIDDEN SHEET

### Table 1 — 'Hidden Rates'!A1:B4

| | A | B |
|---|---|---|
| 1 | Currency | Rate to USD |
| 2 | EUR | 1.08 |
| 3 | GBP | 1.27 |
| 4 | JPY | 0.0067 |

## Sheet 'Config Secret' — VERY HIDDEN SHEET

### Table 1 — 'Config Secret'!A1:B3

| | A | B |
|---|---|---|
| 1 | Setting | Value |
| 2 | tax_rate | 0.21 |
| 3 | fiscal_year_start | April |

## Sheet 'Hidden Rows Cols' — visible

Hidden rows: [3, 5]; hidden columns: ['C'] (shown below and marked)

### Table 1 — 'Hidden Rows Cols'!A1:D5

| | A | B | C (hidden) | D |
|---|---|---|---|---|
| 1 | ID | Name | Salary band | Status |
| 2 | 1 | Ana | B2 | Active |
| 3 (hidden) | 2 | Ben | B3 | Left |
| 4 | 3 | Cy | B1 | Active |
| 5 (hidden) | 4 | Dee | B4 | Left |

## Sheet Empty — visible

(empty sheet)

## Sheet 'Big Log' — visible

### Table 1 — 'Big Log'!A1:D451 (part 1 of 3, rows 1–200)

| | A | B | C | D |
|---|---|---|---|---|
| 1 | Row ID | Event | Value | Flag |
| 2 | 1 | event_1 | 1.5 | FALSE |
| 3 | 2 | event_2 | 3 | FALSE |
| 4 | 3 | event_3 | 4.5 | FALSE |
| 5 | 4 | event_4 | 6 | FALSE |
| 6 | 5 | event_5 | 7.5 | FALSE |
| 7 | 6 | event_6 | 9 | FALSE |
| 8 | 7 | event_7 | 10.5 | TRUE |
| 9 | 8 | event_8 | 12 | FALSE |
| 10 | 9 | event_9 | 13.5 | FALSE |
| 11 | 10 | event_10 | 15 | FALSE |
| 12 | 11 | event_11 | 16.5 | FALSE |
| 13 | 12 | event_12 | 18 | FALSE |
| 14 | 13 | event_13 | 19.5 | FALSE |
| 15 | 14 | event_14 | 21 | TRUE |
| 16 | 15 | event_15 | 22.5 | FALSE |
| 17 | 16 | event_16 | 24 | FALSE |
| 18 | 17 | event_17 | 25.5 | FALSE |
| 19 | 18 | event_18 | 27 | FALSE |
| 20 | 19 | event_19 | 28.5 | FALSE |
| 21 | 20 | event_20 | 30 | FALSE |
| 22 | 21 | event_21 | 31.5 | TRUE |
| 23 | 22 | event_22 | 33 | FALSE |
| 24 | 23 | event_23 | 34.5 | FALSE |
| 25 | 24 | event_24 | 36 | FALSE |
| 26 | 25 | event_25 | 37.5 | FALSE |
| 27 | 26 | event_26 | 39 | FALSE |
| 28 | 27 | event_27 | 40.5 | FALSE |
| 29 | 28 | event_28 | 42 | TRUE |
| 30 | 29 | event_29 | 43.5 | FALSE |
| 31 | 30 | event_30 | 45 | FALSE |
| 32 | 31 | event_31 | 46.5 | FALSE |
| 33 | 32 | event_32 | 48 | FALSE |
| 34 | 33 | event_33 | 49.5 | FALSE |
| 35 | 34 | event_34 | 51 | FALSE |
| 36 | 35 | event_35 | 52.5 | TRUE |
| 37 | 36 | event_36 | 54 | FALSE |
| 38 | 37 | event_37 | 55.5 | FALSE |
| 39 | 38 | event_38 | 57 | FALSE |
| 40 | 39 | event_39 | 58.5 | FALSE |
| 41 | 40 | event_40 | 60 | FALSE |
| 42 | 41 | event_41 | 61.5 | FALSE |
| 43 | 42 | event_42 | 63 | TRUE |
| 44 | 43 | event_43 | 64.5 | FALSE |
| 45 | 44 | event_44 | 66 | FALSE |
| 46 | 45 | event_45 | 67.5 | FALSE |
| 47 | 46 | event_46 | 69 | FALSE |
| 48 | 47 | event_47 | 70.5 | FALSE |
| 49 | 48 | event_48 | 72 | FALSE |
| 50 | 49 | event_49 | 73.5 | TRUE |
| 51 | 50 | event_50 | 75 | FALSE |
| 52 | 51 | event_51 | 76.5 | FALSE |
| 53 | 52 | event_52 | 78 | FALSE |
| 54 | 53 | event_53 | 79.5 | FALSE |
| 55 | 54 | event_54 | 81 | FALSE |
| 56 | 55 | event_55 | 82.5 | FALSE |
| 57 | 56 | event_56 | 84 | TRUE |
| 58 | 57 | event_57 | 85.5 | FALSE |
| 59 | 58 | event_58 | 87 | FALSE |
| 60 | 59 | event_59 | 88.5 | FALSE |
| 61 | 60 | event_60 | 90 | FALSE |
| 62 | 61 | event_61 | 91.5 | FALSE |
| 63 | 62 | event_62 | 93 | FALSE |
| 64 | 63 | event_63 | 94.5 | TRUE |
| 65 | 64 | event_64 | 96 | FALSE |
| 66 | 65 | event_65 | 97.5 | FALSE |
| 67 | 66 | event_66 | 99 | FALSE |
| 68 | 67 | event_67 | 100.5 | FALSE |
| 69 | 68 | event_68 | 102 | FALSE |
| 70 | 69 | event_69 | 103.5 | FALSE |
| 71 | 70 | event_70 | 105 | TRUE |
| 72 | 71 | event_71 | 106.5 | FALSE |
| 73 | 72 | event_72 | 108 | FALSE |
| 74 | 73 | event_73 | 109.5 | FALSE |
| 75 | 74 | event_74 | 111 | FALSE |
| 76 | 75 | event_75 | 112.5 | FALSE |
| 77 | 76 | event_76 | 114 | FALSE |
| 78 | 77 | event_77 | 115.5 | TRUE |
| 79 | 78 | event_78 | 117 | FALSE |
| 80 | 79 | event_79 | 118.5 | FALSE |
| 81 | 80 | event_80 | 120 | FALSE |
| 82 | 81 | event_81 | 121.5 | FALSE |
| 83 | 82 | event_82 | 123 | FALSE |
| 84 | 83 | event_83 | 124.5 | FALSE |
| 85 | 84 | event_84 | 126 | TRUE |
| 86 | 85 | event_85 | 127.5 | FALSE |
| 87 | 86 | event_86 | 129 | FALSE |
| 88 | 87 | event_87 | 130.5 | FALSE |
| 89 | 88 | event_88 | 132 | FALSE |
| 90 | 89 | event_89 | 133.5 | FALSE |
| 91 | 90 | event_90 | 135 | FALSE |
| 92 | 91 | event_91 | 136.5 | TRUE |
| 93 | 92 | event_92 | 138 | FALSE |
| 94 | 93 | event_93 | 139.5 | FALSE |
| 95 | 94 | event_94 | 141 | FALSE |
| 96 | 95 | event_95 | 142.5 | FALSE |
| 97 | 96 | event_96 | 144 | FALSE |
| 98 | 97 | event_97 | 145.5 | FALSE |
| 99 | 98 | event_98 | 147 | TRUE |
| 100 | 99 | event_99 | 148.5 | FALSE |
| 101 | 100 | event_100 | 150 | FALSE |
| 102 | 101 | event_101 | 151.5 | FALSE |
| 103 | 102 | event_102 | 153 | FALSE |
| 104 | 103 | event_103 | 154.5 | FALSE |
| 105 | 104 | event_104 | 156 | FALSE |
| 106 | 105 | event_105 | 157.5 | TRUE |
| 107 | 106 | event_106 | 159 | FALSE |
| 108 | 107 | event_107 | 160.5 | FALSE |
| 109 | 108 | event_108 | 162 | FALSE |
| 110 | 109 | event_109 | 163.5 | FALSE |
| 111 | 110 | event_110 | 165 | FALSE |
| 112 | 111 | event_111 | 166.5 | FALSE |
| 113 | 112 | event_112 | 168 | TRUE |
| 114 | 113 | event_113 | 169.5 | FALSE |
| 115 | 114 | event_114 | 171 | FALSE |
| 116 | 115 | event_115 | 172.5 | FALSE |
| 117 | 116 | event_116 | 174 | FALSE |
| 118 | 117 | event_117 | 175.5 | FALSE |
| 119 | 118 | event_118 | 177 | FALSE |
| 120 | 119 | event_119 | 178.5 | TRUE |
| 121 | 120 | event_120 | 180 | FALSE |
| 122 | 121 | event_121 | 181.5 | FALSE |
| 123 | 122 | event_122 | 183 | FALSE |
| 124 | 123 | event_123 | 184.5 | FALSE |
| 125 | 124 | event_124 | 186 | FALSE |
| 126 | 125 | event_125 | 187.5 | FALSE |
| 127 | 126 | event_126 | 189 | TRUE |
| 128 | 127 | event_127 | 190.5 | FALSE |
| 129 | 128 | event_128 | 192 | FALSE |
| 130 | 129 | event_129 | 193.5 | FALSE |
| 131 | 130 | event_130 | 195 | FALSE |
| 132 | 131 | event_131 | 196.5 | FALSE |
| 133 | 132 | event_132 | 198 | FALSE |
| 134 | 133 | event_133 | 199.5 | TRUE |
| 135 | 134 | event_134 | 201 | FALSE |
| 136 | 135 | event_135 | 202.5 | FALSE |
| 137 | 136 | event_136 | 204 | FALSE |
| 138 | 137 | event_137 | 205.5 | FALSE |
| 139 | 138 | event_138 | 207 | FALSE |
| 140 | 139 | event_139 | 208.5 | FALSE |
| 141 | 140 | event_140 | 210 | TRUE |
| 142 | 141 | event_141 | 211.5 | FALSE |
| 143 | 142 | event_142 | 213 | FALSE |
| 144 | 143 | event_143 | 214.5 | FALSE |
| 145 | 144 | event_144 | 216 | FALSE |
| 146 | 145 | event_145 | 217.5 | FALSE |
| 147 | 146 | event_146 | 219 | FALSE |
| 148 | 147 | event_147 | 220.5 | TRUE |
| 149 | 148 | event_148 | 222 | FALSE |
| 150 | 149 | event_149 | 223.5 | FALSE |
| 151 | 150 | event_150 | 225 | FALSE |
| 152 | 151 | event_151 | 226.5 | FALSE |
| 153 | 152 | event_152 | 228 | FALSE |
| 154 | 153 | event_153 | 229.5 | FALSE |
| 155 | 154 | event_154 | 231 | TRUE |
| 156 | 155 | event_155 | 232.5 | FALSE |
| 157 | 156 | event_156 | 234 | FALSE |
| 158 | 157 | event_157 | 235.5 | FALSE |
| 159 | 158 | event_158 | 237 | FALSE |
| 160 | 159 | event_159 | 238.5 | FALSE |
| 161 | 160 | event_160 | 240 | FALSE |
| 162 | 161 | event_161 | 241.5 | TRUE |
| 163 | 162 | event_162 | 243 | FALSE |
| 164 | 163 | event_163 | 244.5 | FALSE |
| 165 | 164 | event_164 | 246 | FALSE |
| 166 | 165 | event_165 | 247.5 | FALSE |
| 167 | 166 | event_166 | 249 | FALSE |
| 168 | 167 | event_167 | 250.5 | FALSE |
| 169 | 168 | event_168 | 252 | TRUE |
| 170 | 169 | event_169 | 253.5 | FALSE |
| 171 | 170 | event_170 | 255 | FALSE |
| 172 | 171 | event_171 | 256.5 | FALSE |
| 173 | 172 | event_172 | 258 | FALSE |
| 174 | 173 | event_173 | 259.5 | FALSE |
| 175 | 174 | event_174 | 261 | FALSE |
| 176 | 175 | event_175 | 262.5 | TRUE |
| 177 | 176 | event_176 | 264 | FALSE |
| 178 | 177 | event_177 | 265.5 | FALSE |
| 179 | 178 | event_178 | 267 | FALSE |
| 180 | 179 | event_179 | 268.5 | FALSE |
| 181 | 180 | event_180 | 270 | FALSE |
| 182 | 181 | event_181 | 271.5 | FALSE |
| 183 | 182 | event_182 | 273 | TRUE |
| 184 | 183 | event_183 | 274.5 | FALSE |
| 185 | 184 | event_184 | 276 | FALSE |
| 186 | 185 | event_185 | 277.5 | FALSE |
| 187 | 186 | event_186 | 279 | FALSE |
| 188 | 187 | event_187 | 280.5 | FALSE |
| 189 | 188 | event_188 | 282 | FALSE |
| 190 | 189 | event_189 | 283.5 | TRUE |
| 191 | 190 | event_190 | 285 | FALSE |
| 192 | 191 | event_191 | 286.5 | FALSE |
| 193 | 192 | event_192 | 288 | FALSE |
| 194 | 193 | event_193 | 289.5 | FALSE |
| 195 | 194 | event_194 | 291 | FALSE |
| 196 | 195 | event_195 | 292.5 | FALSE |
| 197 | 196 | event_196 | 294 | TRUE |
| 198 | 197 | event_197 | 295.5 | FALSE |
| 199 | 198 | event_198 | 297 | FALSE |
| 200 | 199 | event_199 | 298.5 | FALSE |

### Table 1 — 'Big Log'!A1:D451 (part 2 of 3, rows 201–400)

| | A | B | C | D |
|---|---|---|---|---|
| 1 (first row repeated for context) | Row ID | Event | Value | Flag |
| 201 | 200 | event_200 | 300 | FALSE |
| 202 | 201 | event_201 | 301.5 | FALSE |
| 203 | 202 | event_202 | 303 | FALSE |
| 204 | 203 | event_203 | 304.5 | TRUE |
| 205 | 204 | event_204 | 306 | FALSE |
| 206 | 205 | event_205 | 307.5 | FALSE |
| 207 | 206 | event_206 | 309 | FALSE |
| 208 | 207 | event_207 | 310.5 | FALSE |
| 209 | 208 | event_208 | 312 | FALSE |
| 210 | 209 | event_209 | 313.5 | FALSE |
| 211 | 210 | event_210 | 315 | TRUE |
| 212 | 211 | event_211 | 316.5 | FALSE |
| 213 | 212 | event_212 | 318 | FALSE |
| 214 | 213 | event_213 | 319.5 | FALSE |
| 215 | 214 | event_214 | 321 | FALSE |
| 216 | 215 | event_215 | 322.5 | FALSE |
| 217 | 216 | event_216 | 324 | FALSE |
| 218 | 217 | event_217 | 325.5 | TRUE |
| 219 | 218 | event_218 | 327 | FALSE |
| 220 | 219 | event_219 | 328.5 | FALSE |
| 221 | 220 | event_220 | 330 | FALSE |
| 222 | 221 | event_221 | 331.5 | FALSE |
| 223 | 222 | event_222 | 333 | FALSE |
| 224 | 223 | event_223 | 334.5 | FALSE |
| 225 | 224 | event_224 | 336 | TRUE |
| 226 | 225 | event_225 | 337.5 | FALSE |
| 227 | 226 | event_226 | 339 | FALSE |
| 228 | 227 | event_227 | 340.5 | FALSE |
| 229 | 228 | event_228 | 342 | FALSE |
| 230 | 229 | event_229 | 343.5 | FALSE |
| 231 | 230 | event_230 | 345 | FALSE |
| 232 | 231 | event_231 | 346.5 | TRUE |
| 233 | 232 | event_232 | 348 | FALSE |
| 234 | 233 | event_233 | 349.5 | FALSE |
| 235 | 234 | event_234 | 351 | FALSE |
| 236 | 235 | event_235 | 352.5 | FALSE |
| 237 | 236 | event_236 | 354 | FALSE |
| 238 | 237 | event_237 | 355.5 | FALSE |
| 239 | 238 | event_238 | 357 | TRUE |
| 240 | 239 | event_239 | 358.5 | FALSE |
| 241 | 240 | event_240 | 360 | FALSE |
| 242 | 241 | event_241 | 361.5 | FALSE |
| 243 | 242 | event_242 | 363 | FALSE |
| 244 | 243 | event_243 | 364.5 | FALSE |
| 245 | 244 | event_244 | 366 | FALSE |
| 246 | 245 | event_245 | 367.5 | TRUE |
| 247 | 246 | event_246 | 369 | FALSE |
| 248 | 247 | event_247 | 370.5 | FALSE |
| 249 | 248 | event_248 | 372 | FALSE |
| 250 | 249 | event_249 | 373.5 | FALSE |
| 251 | 250 | event_250 | 375 | FALSE |
| 252 | 251 | event_251 | 376.5 | FALSE |
| 253 | 252 | event_252 | 378 | TRUE |
| 254 | 253 | event_253 | 379.5 | FALSE |
| 255 | 254 | event_254 | 381 | FALSE |
| 256 | 255 | event_255 | 382.5 | FALSE |
| 257 | 256 | event_256 | 384 | FALSE |
| 258 | 257 | event_257 | 385.5 | FALSE |
| 259 | 258 | event_258 | 387 | FALSE |
| 260 | 259 | event_259 | 388.5 | TRUE |
| 261 | 260 | event_260 | 390 | FALSE |
| 262 | 261 | event_261 | 391.5 | FALSE |
| 263 | 262 | event_262 | 393 | FALSE |
| 264 | 263 | event_263 | 394.5 | FALSE |
| 265 | 264 | event_264 | 396 | FALSE |
| 266 | 265 | event_265 | 397.5 | FALSE |
| 267 | 266 | event_266 | 399 | TRUE |
| 268 | 267 | event_267 | 400.5 | FALSE |
| 269 | 268 | event_268 | 402 | FALSE |
| 270 | 269 | event_269 | 403.5 | FALSE |
| 271 | 270 | event_270 | 405 | FALSE |
| 272 | 271 | event_271 | 406.5 | FALSE |
| 273 | 272 | event_272 | 408 | FALSE |
| 274 | 273 | event_273 | 409.5 | TRUE |
| 275 | 274 | event_274 | 411 | FALSE |
| 276 | 275 | event_275 | 412.5 | FALSE |
| 277 | 276 | event_276 | 414 | FALSE |
| 278 | 277 | event_277 | 415.5 | FALSE |
| 279 | 278 | event_278 | 417 | FALSE |
| 280 | 279 | event_279 | 418.5 | FALSE |
| 281 | 280 | event_280 | 420 | TRUE |
| 282 | 281 | event_281 | 421.5 | FALSE |
| 283 | 282 | event_282 | 423 | FALSE |
| 284 | 283 | event_283 | 424.5 | FALSE |
| 285 | 284 | event_284 | 426 | FALSE |
| 286 | 285 | event_285 | 427.5 | FALSE |
| 287 | 286 | event_286 | 429 | FALSE |
| 288 | 287 | event_287 | 430.5 | TRUE |
| 289 | 288 | event_288 | 432 | FALSE |
| 290 | 289 | event_289 | 433.5 | FALSE |
| 291 | 290 | event_290 | 435 | FALSE |
| 292 | 291 | event_291 | 436.5 | FALSE |
| 293 | 292 | event_292 | 438 | FALSE |
| 294 | 293 | event_293 | 439.5 | FALSE |
| 295 | 294 | event_294 | 441 | TRUE |
| 296 | 295 | event_295 | 442.5 | FALSE |
| 297 | 296 | event_296 | 444 | FALSE |
| 298 | 297 | event_297 | 445.5 | FALSE |
| 299 | 298 | event_298 | 447 | FALSE |
| 300 | 299 | event_299 | 448.5 | FALSE |
| 301 | 300 | event_300 | 450 | FALSE |
| 302 | 301 | event_301 | 451.5 | TRUE |
| 303 | 302 | event_302 | 453 | FALSE |
| 304 | 303 | event_303 | 454.5 | FALSE |
| 305 | 304 | event_304 | 456 | FALSE |
| 306 | 305 | event_305 | 457.5 | FALSE |
| 307 | 306 | event_306 | 459 | FALSE |
| 308 | 307 | event_307 | 460.5 | FALSE |
| 309 | 308 | event_308 | 462 | TRUE |
| 310 | 309 | event_309 | 463.5 | FALSE |
| 311 | 310 | event_310 | 465 | FALSE |
| 312 | 311 | event_311 | 466.5 | FALSE |
| 313 | 312 | event_312 | 468 | FALSE |
| 314 | 313 | event_313 | 469.5 | FALSE |
| 315 | 314 | event_314 | 471 | FALSE |
| 316 | 315 | event_315 | 472.5 | TRUE |
| 317 | 316 | event_316 | 474 | FALSE |
| 318 | 317 | event_317 | 475.5 | FALSE |
| 319 | 318 | event_318 | 477 | FALSE |
| 320 | 319 | event_319 | 478.5 | FALSE |
| 321 | 320 | event_320 | 480 | FALSE |
| 322 | 321 | event_321 | 481.5 | FALSE |
| 323 | 322 | event_322 | 483 | TRUE |
| 324 | 323 | event_323 | 484.5 | FALSE |
| 325 | 324 | event_324 | 486 | FALSE |
| 326 | 325 | event_325 | 487.5 | FALSE |
| 327 | 326 | event_326 | 489 | FALSE |
| 328 | 327 | event_327 | 490.5 | FALSE |
| 329 | 328 | event_328 | 492 | FALSE |
| 330 | 329 | event_329 | 493.5 | TRUE |
| 331 | 330 | event_330 | 495 | FALSE |
| 332 | 331 | event_331 | 496.5 | FALSE |
| 333 | 332 | event_332 | 498 | FALSE |
| 334 | 333 | event_333 | 499.5 | FALSE |
| 335 | 334 | event_334 | 501 | FALSE |
| 336 | 335 | event_335 | 502.5 | FALSE |
| 337 | 336 | event_336 | 504 | TRUE |
| 338 | 337 | event_337 | 505.5 | FALSE |
| 339 | 338 | event_338 | 507 | FALSE |
| 340 | 339 | event_339 | 508.5 | FALSE |
| 341 | 340 | event_340 | 510 | FALSE |
| 342 | 341 | event_341 | 511.5 | FALSE |
| 343 | 342 | event_342 | 513 | FALSE |
| 344 | 343 | event_343 | 514.5 | TRUE |
| 345 | 344 | event_344 | 516 | FALSE |
| 346 | 345 | event_345 | 517.5 | FALSE |
| 347 | 346 | event_346 | 519 | FALSE |
| 348 | 347 | event_347 | 520.5 | FALSE |
| 349 | 348 | event_348 | 522 | FALSE |
| 350 | 349 | event_349 | 523.5 | FALSE |
| 351 | 350 | event_350 | 525 | TRUE |
| 352 | 351 | event_351 | 526.5 | FALSE |
| 353 | 352 | event_352 | 528 | FALSE |
| 354 | 353 | event_353 | 529.5 | FALSE |
| 355 | 354 | event_354 | 531 | FALSE |
| 356 | 355 | event_355 | 532.5 | FALSE |
| 357 | 356 | event_356 | 534 | FALSE |
| 358 | 357 | event_357 | 535.5 | TRUE |
| 359 | 358 | event_358 | 537 | FALSE |
| 360 | 359 | event_359 | 538.5 | FALSE |
| 361 | 360 | event_360 | 540 | FALSE |
| 362 | 361 | event_361 | 541.5 | FALSE |
| 363 | 362 | event_362 | 543 | FALSE |
| 364 | 363 | event_363 | 544.5 | FALSE |
| 365 | 364 | event_364 | 546 | TRUE |
| 366 | 365 | event_365 | 547.5 | FALSE |
| 367 | 366 | event_366 | 549 | FALSE |
| 368 | 367 | event_367 | 550.5 | FALSE |
| 369 | 368 | event_368 | 552 | FALSE |
| 370 | 369 | event_369 | 553.5 | FALSE |
| 371 | 370 | event_370 | 555 | FALSE |
| 372 | 371 | event_371 | 556.5 | TRUE |
| 373 | 372 | event_372 | 558 | FALSE |
| 374 | 373 | event_373 | 559.5 | FALSE |
| 375 | 374 | event_374 | 561 | FALSE |
| 376 | 375 | event_375 | 562.5 | FALSE |
| 377 | 376 | event_376 | 564 | FALSE |
| 378 | 377 | event_377 | 565.5 | FALSE |
| 379 | 378 | event_378 | 567 | TRUE |
| 380 | 379 | event_379 | 568.5 | FALSE |
| 381 | 380 | event_380 | 570 | FALSE |
| 382 | 381 | event_381 | 571.5 | FALSE |
| 383 | 382 | event_382 | 573 | FALSE |
| 384 | 383 | event_383 | 574.5 | FALSE |
| 385 | 384 | event_384 | 576 | FALSE |
| 386 | 385 | event_385 | 577.5 | TRUE |
| 387 | 386 | event_386 | 579 | FALSE |
| 388 | 387 | event_387 | 580.5 | FALSE |
| 389 | 388 | event_388 | 582 | FALSE |
| 390 | 389 | event_389 | 583.5 | FALSE |
| 391 | 390 | event_390 | 585 | FALSE |
| 392 | 391 | event_391 | 586.5 | FALSE |
| 393 | 392 | event_392 | 588 | TRUE |
| 394 | 393 | event_393 | 589.5 | FALSE |
| 395 | 394 | event_394 | 591 | FALSE |
| 396 | 395 | event_395 | 592.5 | FALSE |
| 397 | 396 | event_396 | 594 | FALSE |
| 398 | 397 | event_397 | 595.5 | FALSE |
| 399 | 398 | event_398 | 597 | FALSE |
| 400 | 399 | event_399 | 598.5 | TRUE |

### Table 1 — 'Big Log'!A1:D451 (part 3 of 3, rows 401–451)

| | A | B | C | D |
|---|---|---|---|---|
| 1 (first row repeated for context) | Row ID | Event | Value | Flag |
| 401 | 400 | event_400 | 600 | FALSE |
| 402 | 401 | event_401 | 601.5 | FALSE |
| 403 | 402 | event_402 | 603 | FALSE |
| 404 | 403 | event_403 | 604.5 | FALSE |
| 405 | 404 | event_404 | 606 | FALSE |
| 406 | 405 | event_405 | 607.5 | FALSE |
| 407 | 406 | event_406 | 609 | TRUE |
| 408 | 407 | event_407 | 610.5 | FALSE |
| 409 | 408 | event_408 | 612 | FALSE |
| 410 | 409 | event_409 | 613.5 | FALSE |
| 411 | 410 | event_410 | 615 | FALSE |
| 412 | 411 | event_411 | 616.5 | FALSE |
| 413 | 412 | event_412 | 618 | FALSE |
| 414 | 413 | event_413 | 619.5 | TRUE |
| 415 | 414 | event_414 | 621 | FALSE |
| 416 | 415 | event_415 | 622.5 | FALSE |
| 417 | 416 | event_416 | 624 | FALSE |
| 418 | 417 | event_417 | 625.5 | FALSE |
| 419 | 418 | event_418 | 627 | FALSE |
| 420 | 419 | event_419 | 628.5 | FALSE |
| 421 | 420 | event_420 | 630 | TRUE |
| 422 | 421 | event_421 | 631.5 | FALSE |
| 423 | 422 | event_422 | 633 | FALSE |
| 424 | 423 | event_423 | 634.5 | FALSE |
| 425 | 424 | event_424 | 636 | FALSE |
| 426 | 425 | event_425 | 637.5 | FALSE |
| 427 | 426 | event_426 | 639 | FALSE |
| 428 | 427 | event_427 | 640.5 | TRUE |
| 429 | 428 | event_428 | 642 | FALSE |
| 430 | 429 | event_429 | 643.5 | FALSE |
| 431 | 430 | event_430 | 645 | FALSE |
| 432 | 431 | event_431 | 646.5 | FALSE |
| 433 | 432 | event_432 | 648 | FALSE |
| 434 | 433 | event_433 | 649.5 | FALSE |
| 435 | 434 | event_434 | 651 | TRUE |
| 436 | 435 | event_435 | 652.5 | FALSE |
| 437 | 436 | event_436 | 654 | FALSE |
| 438 | 437 | event_437 | 655.5 | FALSE |
| 439 | 438 | event_438 | 657 | FALSE |
| 440 | 439 | event_439 | 658.5 | FALSE |
| 441 | 440 | event_440 | 660 | FALSE |
| 442 | 441 | event_441 | 661.5 | TRUE |
| 443 | 442 | event_442 | 663 | FALSE |
| 444 | 443 | event_443 | 664.5 | FALSE |
| 445 | 444 | event_444 | 666 | FALSE |
| 446 | 445 | event_445 | 667.5 | FALSE |
| 447 | 446 | event_446 | 669 | FALSE |
| 448 | 447 | event_447 | 670.5 | FALSE |
| 449 | 448 | event_448 | 672 | TRUE |
| 450 | 449 | event_449 | 673.5 | FALSE |
| 451 | 450 | event_450 | 675 | FALSE |

## Sheet Messy — visible

Regions on this sheet: cell A1, cell F2, cell M5:R12, cell H8, cell D10, table J30:K31, cell Z50

**A1** (single cell): Top-left note

**F2** (single cell): xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx… [cut: 1500 characters in total]

**M5** (merged cell M5:R12): Big merge value

- Merged cells (value sits in the first cell of the range): M5:R12 ("Big merge value")

**H8** (single cell): #DIV/0!

**D10** (single cell): Line one<br>Line two \| with pipe

### Table 1 — Messy!J30:K31

| | J | K |
|---|---|---|
| 30 | Diagonal A |  |
| 31 |  | Diagonal B |

**Z50** (single cell): Far away cell

## Sheet 'Merged Cells' — visible

Regions on this sheet: cell A1:F1, table A3:F9, table H3:K5, cell A13

**A1** (merged cell A1:F1): Sales by Region and Half-year

- Merged cells (value sits in the first cell of the range): A1:F1 ("Sales by Region and Half-year")

### Table 1 — 'Merged Cells'!A3:F9

| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| 3 | Region | Country | H1 |  | H2 |  |
| 4 |  |  | Q1 | Q2 | Q3 | Q4 |
| 5 | Europe | France | 10 | 12 | 14 | 16 |
| 6 |  | Spain | 8 | 9 | 11 | 13 |
| 7 |  | Italy | 7 | 7 | 9 | 10 |
| 8 | Asia | Japan | 20 | 22 | 21 | 25 |
| 9 |  | India | 15 | 18 | 19 | 24 |

- Merged cells (value sits in the first cell of the range): A3:A4 ("Region"); B3:B4 ("Country"); C3:D3 ("H1"); E3:F3 ("H2"); A5:A7 ("Europe"); A8:A9 ("Asia")

### Table 2 — 'Merged Cells'!H3:K5

| | H | I | J | K |
|---|---|---|---|---|
| 3 | Target band |  | Low | High |
| 4 |  |  | 1 | 2 |
| 5 | Europe | Total | 3 | 4 |

- Merged cells (value sits in the first cell of the range): H3:I4 ("Target band")

**A13** (single cell): After the empty merge

## Sheet Comments — visible

Regions on this sheet: table A1:C4, cell A7, comments on E10

### Table 1 — Comments!A1:C4

| | A | B | C |
|---|---|---|---|
| 1 | Item | Cost | Approved |
| 2 | Pen | 2.5 | Yes |
| 3 | Desk | 180 |  |
| 4 | Lamp | 35 | Yes |

- Comment on B1: Cost in USD, excluding tax.
- Comment on B3: Quote expired. / Ask the vendor for a new one.
- Comment on C3: Waiting for sign-off
- Comment on A4: Comment: Check this figure with Ops / Reply: Confirmed, figure is correct

**A7** (single cell): Legend: Yes = approved

- Comment on A7: Added by audit

### Comments on cells outside any table — Comments

- Comment on E10: Orphan note: ignore the Lamp row

## Sheet Dates — visible

### Table 1 — Dates!A1:C10

| | A | B | C |
|---|---|---|---|
| 1 | Case | Stored value | Excel displays |
| 2 | Date (dd/mm/yyyy) | 2026-10-05 | 05/10/2026 |
| 3 | Date (month and year only) | 2026-10-05 | Oct 2026 |
| 4 | Date and time | 2026-10-05 14:30:00 | 2026-10-05 14:30 |
| 5 | Time of day | 09:15:30 | 09:15:30 |
| 6 | Elapsed time (hours) | 36:30:00 | 36:30 |
| 7 | Date from a formula | 2026-11-04 | 04/11/2026 |
| 8 | Empty date (zero) | 1900-01-00 | 00/01/1900 |
| 9 | Date stored as text | 2026-10-05 | 2026-10-05 |
| 10 | Number, not a date | 46300 | 46300 |

## Sheet 'General Dates' — visible

Regions on this sheet: table A1:G4, table I1:J4, table A7:B8, table A11:B14, table E11:F14

### Table 1 — 'General Dates'!A1:G4

| | A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|---|
| 1 | Order ID | Due Date | Created | Amount Due | Date Code | Updated | Posted |
| 2 | 46300 | 2026-10-05 (serial 46300) | 2026-10-05 14:30:00 (serial 46300.6041666667) | 46300 | 12 | 2026-10-05 (serial 46300) | 2026-10-05 |
| 3 | 46301 | 2026-11-04 (serial 46330) | 2026-10-06 12:00:00 (serial 46301.5) | 46301 | 15 | n/a | 2026-11-04 |
| 4 | 46302 | 2026-12-05 (serial 46361) | 2026-10-07 06:00:00 (serial 46302.25) | 46302 | 7 | 2026-10-08 (serial 46303) | 2026-12-05 |

- Date detection (opt-in): column B ("Due Date"): 3 number(s) stored in General format are shown as dates, with the original number in brackets
- Date detection (opt-in): column C ("Created"): 3 number(s) stored in General format are shown as dates, with the original number in brackets
- Date detection (opt-in): column F ("Updated"): 2 number(s) stored in General format are shown as dates, with the original number in brackets

### Table 2 — 'General Dates'!I1:J4

| | I | J |
|---|---|---|
| 1 | Contract dates |  |
| 2 | Start | End |
| 3 | 2026-10-05 (serial 46300) | 2027-10-05 (serial 46665) |
| 4 | 2026-10-06 (serial 46301) | 2027-10-06 (serial 46666) |

- Merged cells (value sits in the first cell of the range): I1:J1 ("Contract dates")
- Date detection (opt-in): column I ("Contract dates"): 2 number(s) stored in General format are shown as dates, with the original number in brackets
- Date detection (opt-in): column J ("Contract dates"): 2 number(s) stored in General format are shown as dates, with the original number in brackets

### Table 3 — 'General Dates'!A7:B8

| | A | B |
|---|---|---|
| 7 | 46300 | 46301 |
| 8 | 46302 | 46303 |

### Table 4 — 'General Dates'!A11:B14

| | A | B |
|---|---|---|
| 11 | Claim | Due Date |
| 12 | A | TBD |
| 13 | B | 2026-10-05 (serial 46300) |
| 14 | C | 2026-10-06 (serial 46301) |

- Date detection (opt-in): column B ("Due Date"): 2 number(s) stored in General format are shown as dates, with the original number in brackets

### Table 5 — 'General Dates'!E11:F14

| | E | F |
|---|---|---|
| 11 | Claim | Dates |
| 12 |  | Start |
| 13 | B | 2026-10-05 (serial 46300) |
| 14 | C | 2026-10-06 (serial 46301) |

- Date detection (opt-in): column F ("Dates"): 2 number(s) stored in General format are shown as dates, with the original number in brackets
