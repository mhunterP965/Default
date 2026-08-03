# Figure 5 (A0) — extracted authoritative topology  [source: ATMOS_IRS_W911W625CA002.docx, rId22 = word/media/image6.png]

Border 33,52 -> 1022,768 px @96dpi. Title centered bold top. Node id "A0" lower-right.
Function-box node id rendered bold in the box's BOTTOM-RIGHT corner.

Boxes(px): A1 109,354,240,426 | A2 268,354,398,426 | A3 426,354,557,426
           A4 580,354,730,426 | A5 791,226,936,298 | A6 791,504,936,576

## Control allocation (from bus/stem tracing)
C1 bus y=142 (stem x=200)  -> A5, A6
C2 bus y=158 (stem x=344)  -> A4, A5, A6
C3 y=182     (stem x=578)  -> A4
C4 bus y=195 (stem x=868)  -> A2 only      [IER-26 => A2 only]

## Mechanism allocation (bus y=642 M1, y=676 M2, y=710 M3)
M1 -> A1,A2,A3,A4,A5,A6      (every box)
M2 -> A2,A3 only
M3 -> A1,A4,A5,A6

## Inputs / flows / outputs
I1,I2 -> A1 left (2 separate ports, y=375 / y=405)
I3    -> A5 left (lane y=220 across, NOT a control)
I4    -> A2 left (y=514 lane, up x=248, into A2 left y=408)
F1 A1->A2 | F2 A2->A3 | F3 A3->A4 | F4 A4->A5 | F5 A4->A6   (all left-face entries at y~390)
O1 <- A4 right y=390 | O3 <- A5 right y=262 | O2 <- A6 right y=540

## Note
Original breaks connector lines behind white label fills (e.g. x=618 under the F3 label).
The replacement must NOT do this (labels may not sit on connectors).
