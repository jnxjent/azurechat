type Rect = { x: number; y: number; width: number; height: number };

/** Recognize framed contact sheets conservatively; ordinary pages stay whole. */
export function findPdfPagePanels(data: Uint8ClampedArray, width: number, height: number): Rect[] {
  const ink = (x: number, y: number) => {
    const i = (y * width + x) * 4;
    return data[i + 3] > 0 && Math.min(data[i], data[i + 1], data[i + 2]) < 250;
  };
  const xs = new Array(width).fill(0), ys = new Array(height).fill(0);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) if (ink(x, y)) { xs[x]++; ys[y]++; }
  const bands = (counts: number[], minimumGap: number, noise: number) => {
    const ranges: Array<[number, number]> = [];
    let start = -1, last = -1;
    counts.forEach((n, i) => {
      if (n > noise) {
        if (start < 0) start = i;
        else if (i - last > minimumGap) { ranges.push([start, last + 1]); start = i; }
        last = i;
      }
    });
    if (start >= 0) ranges.push([start, last + 1]);
    return ranges;
  };
  const columns = bands(xs, 1, height * .03).filter(([a,b]) => b-a > width * .08);
  const edgeCounts = new Array(height).fill(0);
  for (let y = 0; y < height; y++) for (const [a,b] of columns) {
    for (const x of [a,b-1]) {
      if ([-2,-1,0,1,2].some(d => x+d >= 0 && x+d < width && ink(x+d,y))) edgeCounts[y]++;
    }
  }
  const rows = bands(edgeCounts, Math.max(2, Math.round(height * .002)), columns.length * 1.5 - 1)
    .filter(([a,b]) => b-a > height * .04);
  const whole = [{ x: 0, y: 0, width, height }];
  if (columns.length < 2 || columns.length > 4 || rows.length < 2 || rows.length > 4) return whole;
  const rects = rows.flatMap(([y, bottom]) => columns.map(([x, right]) => ({ x, y, width: right - x, height: bottom - y })));
  if (rects.some(r => r.width < 80 || r.height < 80 || r.width / r.height < .4 || r.width / r.height > 2.4)) return whole;
  // A framed slide has ink along all four outside edges. This avoids splitting
  // tables, ordinary multi-column text, and unframed photographs into slides.
  for (const r of rects) {
    const edge = (horizontal: boolean, end: boolean) => {
      const length = horizontal ? r.width : r.height;
      let hits = 0;
      for (let i = 0; i < length; i++) {
        let hit = false;
        for (let d = 0; d < 3; d++) {
          const x = horizontal ? r.x + i : end ? r.x + r.width - 1 - d : r.x + d;
          const y = horizontal ? end ? r.y + r.height - 1 - d : r.y + d : r.y + i;
          hit ||= ink(x, y);
        }
        if (hit) hits++;
      }
      return hits / length;
    };
    const edges = [edge(true, false), edge(true, true), edge(false, false), edge(false, true)];
    if (edges.some(v => v < .6)) return whole;
  }
  return rects;
}
