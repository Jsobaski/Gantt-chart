"use strict";
const { PNG } = require("pngjs");
const fs = require("fs");

const W = 20, H = 20;
const png = new PNG({ width: W, height: H });

const px = (x, y, r, g, b, a = 255) => {
    if (x < 0 || x >= W || y < 0 || y >= H) return;
    const i = (y * W + x) * 4;
    png.data[i] = r; png.data[i+1] = g; png.data[i+2] = b; png.data[i+3] = a;
};

const fill = (x, y, w, h, r, g, b, a = 255) => {
    for (let dy = 0; dy < h; dy++)
        for (let dx = 0; dx < w; dx++)
            px(x + dx, y + dy, r, g, b, a);
};

// Background: dark navy
fill(0, 0, W, H, 13, 27, 42);

// Label column (4px wide, slightly lighter)
fill(0, 0, 4, H, 10, 22, 38);

// Vertical divider between label and chart
fill(4, 0, 1, H, 26, 58, 92);

// Month grid lines (every 5px in the chart area)
for (let gx = 9; gx < W; gx += 5) fill(gx, 0, 1, H, 26, 58, 92, 140);

// ── Location header rows ─────────────────────────────────
fill(0, 0, W, 3, 17, 40, 64);   // Location 1 header
fill(0, 11, W, 3, 17, 40, 64);  // Location 2 header

// ── Project rows ─────────────────────────────────────────
// Row 1 (y=3..6): label stub + two bars + milestone
fill(0, 3, 4, 4, 8, 18, 34);    // label bg
fill(5, 4, 8, 1, 77, 171, 247); // P6 bar (blue)
fill(6, 5, 7, 1, 38, 198, 218); // Maximo bar (teal)
// Milestone diamond at x=5 midpoint y=4
px(5, 3,  171, 71, 188);
px(4, 4,  171, 71, 188); px(5, 4, 171, 71, 188); px(6, 4, 171, 71, 188);
px(5, 5,  171, 71, 188);

// Row 2 (y=7..10): label stub + bars
fill(0, 7, 4, 4, 8, 18, 34);
fill(7, 8, 6, 1, 77, 171, 247);
fill(6, 9, 7, 1, 38, 198, 218);
// green milestone
px(7, 7,  102, 187, 106);
px(6, 8,  102, 187, 106); px(7, 8, 102, 187, 106); px(8, 8, 102, 187, 106);
px(7, 9,  102, 187, 106);

// Row 3 (y=14..18): label stub + bars
fill(0, 14, 4, 5, 8, 18, 34);
fill(5, 15, 10, 1, 77, 171, 247);
fill(6, 16, 9, 1, 38, 198, 218);
// orange milestone
px(10, 14, 255, 167, 38);
px(9,  15, 255, 167, 38); px(10, 15, 255, 167, 38); px(11, 15, 255, 167, 38);
px(10, 16, 255, 167, 38);

// Row 4 (y=18..19): another small row
fill(0, 18, 4, 2, 8, 18, 34);
fill(8, 18, 5, 1, 77, 171, 247);
fill(7, 19, 6, 1, 38, 198, 218);

fs.writeFileSync("assets/icon.png", PNG.sync.write(png));
console.log("Icon generated: assets/icon.png");
