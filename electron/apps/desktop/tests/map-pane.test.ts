import { describe, expect, it } from "vitest";
import { projectWorldPoint } from "../src/renderer/src/lib/map-projection.js";

describe("projectWorldPoint", () => {
  it("keeps Rust world coordinates inside the centered playable map square", () => {
    const map = { width: 6000, height: 6000, worldSize: 4000, oceanMargin: 1000, imageBase64: null, monuments: [] };
    expect(projectWorldPoint(map, 0, 0)).toEqual({ left: 16.666666666666664, top: 83.33333333333334 });
    expect(projectWorldPoint(map, 2000, 2000)).toEqual({ left: 50, top: 50 });
    expect(projectWorldPoint(map, 4000, 4000)).toEqual({ left: 83.33333333333334, top: 16.666666666666664 });
  });
});
