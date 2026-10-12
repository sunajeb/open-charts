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
  combo: {
    specVersion: 1,
    type: 'combo',
    categories: years,
    series: [
      { name: 'Products', values: [62, 71, 84, 98, 112] },
      { name: 'Services', values: [24, 27, 33, 38, 44] },
      { name: 'Plan', values: [90, 96, 108, 118, 126], as: 'line' },
    ],
    options: { unit: 'm', totals: true },
  },
  mekkoxy: {
    specVersion: 1,
    type: 'mekkoxy',
    categories: ['Plant A', 'Plant B', 'Plant C', 'Plant D', 'Plant E'],
    series: [
      { name: 'Capacity (Mt)', values: [12, 8, 15, 5, 9] },
      { name: 'Cash cost ($/t)', values: [41, 55, 38, 72, 49] },
    ],
    options: { sort: 'ascending', unit: { text: '$', position: 'prefix' } },
  },
  scatter: {
    specVersion: 1,
    type: 'scatter',
    categories: ['Acme', 'Globex', 'Initech', 'Umbrella', 'Hooli', 'Stark'],
    series: [{ name: 'Accounts', values: [[12.5, 42], [9.8, 39], [4.2, 31], [14.1, 47], [5.6, 33], [11.2, 44]] }],
    options: { axisTitles: { x: 'Seats (thousands)', y: 'Price per seat ($)' } },
  },
  bubble: {
    specVersion: 1,
    type: 'bubble',
    categories: ['Payments', 'Lending', 'Wealth', 'Insurance', 'Cards'],
    series: [{ name: 'Businesses', values: [[8.5, 2.1, 420], [3.2, 1.4, 260], [6.1, 0.6, 180], [1.8, 0.4, 90], [11.4, 1.8, 310]] }],
    options: { zeroBased: true, axisTitles: { x: 'Market growth (%)', y: 'Relative market share (x)' } },
  },
  table: {
    specVersion: 1,
    type: 'table',
    categories: years,
    series: [
      { name: 'Revenue ($m)', values: [91, 120, 148, 181, 216] },
      { name: 'EBITDA ($m)', values: [12, 19, 26, 35, 44] },
      { name: 'EBITDA margin', values: ['13.2%', '15.8%', '17.6%', '19.3%', '20.4%'] },
    ],
  },
  gantt: {
    specVersion: 1,
    type: 'gantt',
    categories: ['Diagnostic', 'Design', 'Pilot', 'Roll-out', 'Steering committee'],
    series: [
      { name: 'Start', values: ['2026-01-05', '2026-02-02', '2026-03-02', '2026-05-04', null] },
      { name: 'End', values: ['2026-02-02', '2026-03-30', '2026-05-04', '2026-11-30', null] },
      { name: 'Milestone', values: ['2026-02-02', null, '2026-05-04', '2026-11-30', '2026-03-16'] },
    ],
  },
  area: {
    specVersion: 1,
    type: 'area',
    categories: years,
    series: [
      { name: 'Enterprise', values: [42, 58, 75, 96, 118] },
      { name: 'Mid-market', values: [31, 36, 44, 51, 60] },
      { name: 'SMB', values: [18, 21, 22, 26, 29] },
    ],
    options: { unit: 'm', totals: true },
  },
  area100: {
    specVersion: 1,
    type: 'area100',
    categories: years,
    series: [
      { name: 'Enterprise', values: [42, 58, 75, 96, 118] },
      { name: 'Mid-market', values: [31, 36, 44, 51, 60] },
      { name: 'SMB', values: [18, 21, 22, 26, 29] },
    ],
  },
  pie: {
    specVersion: 1,
    type: 'pie',
    categories: ['Enterprise', 'Mid-market', 'SMB', 'Public sector', 'Other'],
    series: [{ name: 'Revenue 2025', values: [118, 60, 29, 14, 6] }],
  },
  doughnut: {
    specVersion: 1,
    type: 'doughnut',
    categories: ['Enterprise', 'Mid-market', 'SMB', 'Public sector', 'Other'],
    series: [{ name: 'Revenue 2025', values: [118, 60, 29, 14, 6] }],
    options: { unit: 'm', totals: true },
  },
  harvey: {
    specVersion: 1,
    type: 'harvey',
    categories: ['Price', 'Features', 'Support', 'Integrations'],
    series: [
      { name: 'Us', values: [3, 4, 2, 3] },
      { name: 'Competitor A', values: [4, 2, 3, 1] },
      { name: 'Competitor B', values: [2, 3, 1, 0] },
    ],
  },
};
