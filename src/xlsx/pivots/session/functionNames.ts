/** A data field's default caption prefix by its summary function, as Excel names it ("Sum of Sales"): the menus'
 * labels and GETPIVOTDATA's data-field names read the same table. */
export const PIVOT_FUNCTION_NAMES: Readonly<Record<string, string>> = Object.freeze({
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
})
