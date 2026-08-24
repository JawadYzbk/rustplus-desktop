import type { MapSnapshot } from "../stores/connection.js";

export type MapPoint = { left: number; top: number };

function clamp(value: number): number {
  return Math.max(0, Math.min(100, value));
}

/** Matches MainWindow.WorldToImagePx: coordinates are 0..worldSize, Y is north-up. */
export function projectWorldPoint(map: MapSnapshot, x: number, y: number): MapPoint | null {
  const worldSize = map.worldSize > 0 ? map.worldSize : Math.max(map.width, map.height);
  const imageWidth = map.width > 0 ? map.width : worldSize;
  const imageHeight = map.height > 0 ? map.height : worldSize;
  if (!(worldSize > 0 && imageWidth > 0 && imageHeight > 0) || !Number.isFinite(x) || !Number.isFinite(y)) return null;

  const padWorld = map.oceanMargin > 0 ? map.oceanMargin * 2 : 2000;
  const side = Math.min(imageWidth, imageHeight) * (worldSize / (worldSize + padWorld));
  const offsetX = (imageWidth - side) / 2;
  const offsetY = (imageHeight - side) / 2;
  return {
    left: clamp(((offsetX + (clamp(x / worldSize) * side)) / imageWidth) * 100),
    top: clamp(((offsetY + ((1 - clamp(y / worldSize)) * side)) / imageHeight) * 100),
  };
}
