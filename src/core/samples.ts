// Default sample data per chart type: what the PowerPoint add-in inserts when the user picks a type
// from the ribbon (the reference product seeds its datasheet the same way). Chart-type authors add
// their entry here; a type with an entry appears in the add-in once the engine can lay it out.
import type { ChartSpec, ChartType } from './types.ts';

const years = ['2021', '2022', '2023', '2024', '2025'];

export const samples: Partial<Record<ChartType, ChartSpec>> = {
  stacked: {
    specVersion: 1,
    type: 'stacked',
    categories: years,
    series: [
      { name: 'Enterprise', values: [42, 58, 75, 96, 118] },
      { name: 'Mid-market', values: [31, 36, 44, 51, 60] },
      { name: 'SMB', values: [18, 21, 22, 26, 29] },
    ],
    options: { unit: 'm', cagr: { from: 0, to: 4 } },
  },
  clustered: {
    specVersion: 1,
    type: 'clustered',
    categories: ['North America', 'Europe', 'Asia Pacific', 'Latin America'],
    series: [
      { name: 'Enterprise', values: [38, 31, 27, 22] },
      { name: 'Mid-market', values: [29, 26, 24, 19] },
    ],
    options: { unit: '%' },
  },
  percent: {
    specVersion: 1,
    type: 'percent',
    categories: years,
    series: [
      { name: 'Enterprise', values: [42, 58, 75, 96, 118] },
      { name: 'Mid-market', values: [31, 36, 44, 51, 60] },
      { name: 'SMB', values: [18, 21, 22, 26, 29] },
    ],
  },
  waterfall: {
    specVersion: 1,
    type: 'waterfall',
    categories: ['EBITDA 2024', 'Volume', 'Price', 'Mix', 'Raw materials', 'Labour', 'EBITDA 2025'],
    series: [{ name: 'EBITDA', values: [100, 18, 9, -4, -15, -7, null] }],
    options: { unit: 'm' },
  },
  mekko: {
    specVersion: 1,
    type: 'mekko',
    categories: ['Enterprise', 'Mid-market', 'SMB', 'Public sector'],
    series: [
      { name: 'Us', values: [410, 180, 90, 60] },
      { name: 'Competitor A', values: [290, 240, 160, 70] },
      { name: 'Others', values: [120, 130, 240, 130] },
    ],
  },
  line: {
    specVersion: 1,
    type: 'line',
    categories: years,
    series: [
      { name: 'Revenue', values: [91, 120, 148, 181, 216] },
      { name: 'Costs', values: [84, 101, 119, 138, 152] },
    ],
    options: { unit: 'm' },
  },
};
