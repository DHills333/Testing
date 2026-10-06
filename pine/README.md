# Swing Segment Boxes (TradingView / Pine Script v6)

Draws a shaded box over each leg between 4 to 8 swing points, with a diagonal line through it:
green for up legs, red for down legs. Labels show the price change, % change and bar count.

## Install
1. In TradingView, open the **Pine Editor**, paste in `swing_segment_boxes.pine`, and click **Add to chart**.
2. TradingView then asks you to click **Point 1** to **Point 4** on the chart, e.g. swing low, high, low, high. The points are sorted by time, so click order does not affect the colours.
3. To adjust afterwards, select the indicator and drag the point handles, or edit the times/prices in its settings.

## Adding more points (5 to 8)
TradingView asks for every click-to-place point as soon as the indicator is added, so points 5 to 8 can't be optional clicks. To use them:
1. Open the indicator's settings and raise **Number of points** (4 to 8).
2. Under **Extra points**, set the date/time of each new point.
3. Leave its price at 0 to place it automatically on that bar's low after an up leg, or its high after a down leg. Or type in a price.

## Why not the Polyline tool?
Pine Script can't read drawings you made with TradingView's own tools (Polyline, Path, Trend Line…).
Interactive `input.time` / `input.price` points are the closest supported equivalent: you still click the swing points on the chart, and the indicator draws the boxes.

## Options
- **Snap points to bar high/low**: moves each point to the exact wick of the bar you clicked (high for swing highs, low for swing lows).
- Fill, border and line colours for up and down legs.
- Toggle the diagonal lines and change labels.
