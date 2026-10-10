/** SpreadsheetML ST_Xstring escapes encode UTF-16 units, including units
 * forbidden in XML. A single pass also handles _x005F_x0041_: its decoded
 * leading underscore must not start a second escape-decoding pass. */
export function decodeXlsxString(value:string):string {
  return value.replace(/_x([0-9A-Fa-f]{4})_/g,(_match,hex:string)=>String.fromCharCode(parseInt(hex,16)))
}
