# Trend Segment Boxes (TradingView / Pine Script v6)

Draws a shaded box over each leg between four swing points (three legs), with a diagonal line through it:
green for up legs, red for down legs. Labels show the price change, % change and candle count.

## Install
1. In TradingView, open the **Pine Editor**, paste in `trend_segment_boxes.pine`, and click **Add to chart**.
2. TradingView then asks you to click **Point 1** to **Point 4** on the chart, e.g. swing low, high, low, high. The points are sorted by time, so click order does not affect the colors.
3. TradingView also asks you to confirm **Show only on symbol**. Set it to the symbol you placed the points on (for example `NASDAQ:AAPL`). The indicator stays on the chart when you switch symbols, because TradingView keeps indicators per chart layout, but it draws nothing on other symbols. Leave it empty to draw on every symbol.
4. To adjust afterwards, select the indicator and drag the point handles, or edit the times/prices in its settings.

## Why not the Polyline tool?
Pine Script can't read drawings you made with TradingView's own tools (Polyline, Path, Trend Line…).
Interactive `input.time` / `input.price` points are the closest supported equivalent: you still click the swing points on the chart, and the indicator draws the boxes.

To chain more legs, add the indicator again and start its Point 1 where the previous one's Point 4 ended.

## Options
- **Snap points to bar high/low**: moves each point to the exact wick of the bar you clicked (high for swing highs, low for swing lows).
- Fill, border and line colors for up and down legs.
- Toggle the diagonal lines and change labels, and set the change label size (tiny to huge).

## Controlling swing
The indicator marks the controlling swing with a dashed blue line:
- **Uptrend** (the latest structure break is a higher high): the controlling swing low (**CSL**) is the swing low that launched that higher high.
- **Downtrend** (the latest structure break is a lower low): the controlling swing high (**CSH**) is the swing high that launched that lower low.

The line extends right until a candle closes beyond it. The line then stops at that candle and the label adds "broken". If the four points make no higher high or lower low (a range), nothing is marked.

It also marks the **prior opposite swing** with a dashed blue line: the swing just before the controlling one, on the other side. That's the prior swing high (**PSH**) for a CSL, or the prior swing low (**PSL**) for a CSH. It's the level the break of structure went through, and the line extends right.

Each label sits on its line just to the right of the swing point: under the line for a low, over it for a high.

Under **Controlling swing** in the settings you can turn either line off and set its color, style (solid, dashed or dotted) and width. You can also set the swing label size there.
