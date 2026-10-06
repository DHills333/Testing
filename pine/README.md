# Trend Segment Boxes (TradingView / Pine Script v6)

Draws a shaded box over each leg between four swing points (three legs), with a diagonal line through it:
green for up legs, red for down legs. Labels show the price change, % change and candle count.

## Install
1. In TradingView, open the **Pine Editor**, paste in `trend_segment_boxes.pine`, and click **Add to chart**.
2. TradingView then asks you to click **Point 1** to **Point 4** on the chart, e.g. swing low, high, low, high. The points are sorted by time, so click order does not affect the colours.
3. To adjust afterwards, select the indicator and drag the point handles, or edit the times/prices in its settings.

## Why not the Polyline tool?
Pine Script can't read drawings you made with TradingView's own tools (Polyline, Path, Trend Line…).
Interactive `input.time` / `input.price` points are the closest supported equivalent: you still click the swing points on the chart, and the indicator draws the boxes.

To chain more legs, add the indicator again and start its Point 1 where the previous one's Point 4 ended.

## Options
- **Snap points to bar high/low**: moves each point to the exact wick of the bar you clicked (high for swing highs, low for swing lows).
- Fill, border and line colours for up and down legs.
- Toggle the diagonal lines and change labels.

## Controlling swing
The indicator marks the controlling swing with a dashed line and label:
- **Uptrend** (the latest structure break is a higher high): the **Controlling Low** is the swing low that launched that higher high.
- **Downtrend** (the latest structure break is a lower low): the **Controlling High** is the swing high that launched that lower low.

The line extends right until a candle closes beyond it. The line then stops at that candle and the label says "(broken)". If the four points make no higher high or lower low (a range), nothing is marked. Turn it off or change its colour under **Controlling swing** in the settings.
