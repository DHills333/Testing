Trend Segment Boxes is a manual market-structure tool. You click four swing points on the chart. The script turns the three legs between them into measured boxes, then marks the controlling swing and the prior opposite swing implied by those points.

It doesn't detect swings automatically and it doesn't generate buy or sell signals. Its purpose is to make the structure you have identified easier to read and measure.

HOW IT WORKS

Swing points
• When the indicator is added, TradingView asks you to click Point 1 to Point 4 on the chart. These are interactive time/price inputs, so you can later drag them on the chart or edit them in the settings.
• Points are sorted by time, so the result doesn't depend on the order you click them in.
• You are also asked to confirm "Show only on symbol". The drawings appear only on that symbol, so switching the chart to another symbol doesn't carry your swing points over to it. Leave it empty to draw on every symbol.
• With "Snap points to bar high/low" enabled (the default), each point moves to the exact wick of the candle you clicked. It uses the high if the point is above its neighbors (a swing high) and the low if it is below them (a swing low).

Leg boxes
• Each leg between two consecutive points is drawn as a box spanning the leg's price range and time.
• Rising legs use the "Up" colors and falling legs use the "Down" colors. A diagonal line connects the two swing points.
• A label on each leg shows the price change, the percentage change and the number of candles in the leg. It is centered on the leg, above the box for up legs and below it for down legs.

Controlling swing (CSH / CSL)
• The script looks back from the newest point for the most recent break of structure among your four points. That is a swing high above the previous swing high (a higher high) or a swing low below the previous swing low (a lower low).
• After a higher high, the swing low that launched it is marked CSL (controlling swing low). After a lower low, the swing high that launched it is marked CSH (controlling swing high).
• A horizontal line extends right from that swing. If a later candle closes beyond it, the line stops at that candle and the label adds "broken". A close is used, not a wick. The check covers up to the last 5,000 candles.
• If the four points contain no higher high or lower low (for example a range), nothing is marked.

Prior opposite swing (PSH / PSL)
• This is the swing just before the controlling swing, on the other side. It is the prior swing high (PSH) for a CSL, or the prior swing low (PSL) for a CSH. It is the level the break of structure went through.
• It is drawn as a horizontal line extending right from that swing.

HOW TO USE IT

1. Add the indicator and click four swing points in sequence, for example low, high, low, high.
2. Read the leg labels to compare the size and duration of impulse and pullback legs.
3. Use the CSL/CSH line as the level whose loss would invalidate the current structure, and the PSH/PSL line as the level that was broken to confirm it.
4. When structure develops further, drag the points to the newest swings, or add a second instance of the indicator for the next set of legs.

SETTINGS

• Style: snap to wicks; fill, border and line colors for up and down legs; show/hide diagonal lines and change labels; change label size.
• Controlling swing: show/hide the CSH/CSL and the PSH/PSL lines; color, style (solid, dashed or dotted) and width of each line; swing label size.

NOTES AND LIMITATIONS

• Everything drawn depends on the points you choose. Different swing choices produce different boxes and levels.
• The script works on any symbol and timeframe. Points snap to the candle under the selected time, so changing the timeframe can move a point to a different candle.
• The script can't read drawings made with TradingView's own drawing tools. Its points are set only through its own inputs.

This tool is for chart analysis and education. It doesn't predict future price movement and isn't financial advice.
