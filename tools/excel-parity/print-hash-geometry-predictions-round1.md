# Locked predictions: hash geometry and 7->8 transition

Written before the Excel export; model = tool docstring (W=5.25w+3.75 pt, 5 px padding each side, per-glyph advance quantised at 600 dpi, n=floor(content/advance)).

## Primary invariants (no constants fitted to these sheets)

| alignment | slope d(xMax-left)/dW | slope d(xMin-left)/dW | slope d(right gap)/dW | transitions |
|---|---|---|---|---|
| right | 1.00 | 1.00 | 0.00 | xMax fixed, xMin drops one advance |
| center | 0.50 | 0.50 | 0.50 | xMin -a/2, xMax +a/2 |
| left | 0.00 | 0.00 | 1.00 | xMin fixed, xMax jumps one advance |

Italic = pure translation by the same-sheet non-overflow italic/upright delta; run width unchanged; literal text = generated overflow word box (upright); 75% keeps counts vs widthChars, scales all PDF points by 0.75.

## Predicted counts per width

- **R_aptos14_upright** (overflow, Aptos Narrow 14pt, right, 100%): 8:5, 8.25:5, 8.5:5, 8.75:6, 9:6, 9.25:6, 9.5:6, 9.75:6, 10:7, 10.25:7, 10.5:7, 10.75:7, 11:7, 11.25:7, 11.5:8, 11.75:8, 12:8, 12.25:8, 12.5:8, 12.75:9, 13:9
- **L_aptos14_upright** (overflow, Aptos Narrow 14pt, left, 100%): 8:5, 8.25:5, 8.5:5, 8.75:6, 9:6, 9.25:6, 9.5:6, 9.75:6, 10:7, 10.25:7, 10.5:7, 10.75:7, 11:7, 11.25:7, 11.5:8, 11.75:8, 12:8, 12.25:8, 12.5:8, 12.75:9, 13:9
- **C_aptos14_upright** (overflow, Aptos Narrow 14pt, center, 100%): 8:5, 8.25:5, 8.5:5, 8.75:6, 9:6, 9.25:6, 9.5:6, 9.75:6, 10:7, 10.25:7, 10.5:7, 10.75:7, 11:7, 11.25:7, 11.5:8, 11.75:8, 12:8, 12.25:8, 12.5:8, 12.75:9, 13:9
- **R_aptos14_italic** (overflow, Aptos Narrow 14pt italic, right, 100%): 8:5, 8.25:5, 8.5:5, 8.75:6, 9:6, 9.25:6, 9.5:6, 9.75:6, 10:7, 10.25:7, 10.5:7, 10.75:7, 11:7, 11.25:7, 11.5:8, 11.75:8, 12:8, 12.25:8, 12.5:8, 12.75:9, 13:9
- **L_aptos14_italic** (overflow, Aptos Narrow 14pt italic, left, 100%): 8:5, 8.25:5, 8.5:5, 8.75:6, 9:6, 9.25:6, 9.5:6, 9.75:6, 10:7, 10.25:7, 10.5:7, 10.75:7, 11:7, 11.25:7, 11.5:8, 11.75:8, 12:8, 12.25:8, 12.5:8, 12.75:9, 13:9
- **C_aptos14_italic** (overflow, Aptos Narrow 14pt italic, center, 100%): 8:5, 8.25:5, 8.5:5, 8.75:6, 9:6, 9.25:6, 9.5:6, 9.75:6, 10:7, 10.25:7, 10.5:7, 10.75:7, 11:7, 11.25:7, 11.5:8, 11.75:8, 12:8, 12.25:8, 12.5:8, 12.75:9, 13:9
- **R_literal_aptos14** (literal, Aptos Narrow 14pt, right, 100%): 10:7, 10.5:7, 11:7, 11.5:8, 12:8, 12.5:8, 13:9, 13.5:9, 14:10, 14.5:10, 15:10
- **L_literal_aptos14** (literal, Aptos Narrow 14pt, left, 100%): 10:7, 10.5:7, 11:7, 11.5:8, 12:8, 12.5:8, 13:9, 13.5:9, 14:10, 14.5:10, 15:10
- **C_literal_aptos14** (literal, Aptos Narrow 14pt, center, 100%): 10:7, 10.5:7, 11:7, 11.5:8, 12:8, 12.5:8, 13:9, 13.5:9, 14:10, 14.5:10, 15:10
- **R_literal_aptos14_i** (literal, Aptos Narrow 14pt italic, right, 100%): 10:7, 10.5:7, 11:7, 11.5:8, 12:8, 12.5:8, 13:9, 13.5:9, 14:10, 14.5:10, 15:10
- **R_calibri11** (overflow, Calibri 11pt, right, 100%): 8:6, 8.25:7, 8.5:7, 8.75:7, 9:7, 9.25:8, 9.5:8, 9.75:8, 10:8, 10.25:9, 10.5:9, 10.75:9, 11:9, 11.25:10, 11.5:10, 11.75:10, 12:10, 12.25:10, 12.5:11, 12.75:11, 13:11
- **R_calibri11_italic** (overflow, Calibri 11pt italic, right, 100%): 8:6, 8.25:7, 8.5:7, 8.75:7, 9:7, 9.25:8, 9.5:8, 9.75:8, 10:8, 10.25:9, 10.5:9, 10.75:9, 11:9, 11.25:10, 11.5:10, 11.75:10, 12:10, 12.25:10, 12.5:11, 12.75:11, 13:11
- **R_arial12** (overflow, Arial 12pt, right, 100%): 10:7, 10.25:7, 10.5:7, 10.75:7, 11:8, 11.25:8, 11.5:8, 11.75:8, 12:8, 12.25:9, 12.5:9, 12.75:9, 13:9, 13.25:9, 13.5:9, 13.75:10, 14:10, 14.25:10, 14.5:10, 14.75:10, 15:11, 15.25:11, 15.5:11, 15.75:11, 16:11
- **R_aptos14_75** (overflow, Aptos Narrow 14pt, right, 75%): 8:5, 8.25:5, 8.5:5, 8.75:6, 9:6, 9.25:6, 9.5:6, 9.75:6, 10:7, 10.25:7, 10.5:7, 10.75:7, 11:7, 11.25:7, 11.5:8, 11.75:8, 12:8, 12.25:8, 12.5:8, 12.75:9, 13:9
- **R_aptos14_italic_75** (overflow, Aptos Narrow 14pt italic, right, 75%): 8:5, 8.25:5, 8.5:5, 8.75:6, 9:6, 9.25:6, 9.5:6, 9.75:6, 10:7, 10.25:7, 10.5:7, 10.75:7, 11:7, 11.25:7, 11.5:8, 11.75:8, 12:8, 12.25:8, 12.5:8, 12.75:9, 13:9
