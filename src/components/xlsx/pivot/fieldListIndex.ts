import type { PivotMenuField, PivotMenuModel } from './pivotClient'

/** Shared lookups for rendering and moving fields; construction is one pass over fields and data. */
export interface PivotFieldListIndex {
  fieldsByIndex: ReadonlyMap<number, PivotMenuField>
  dataIndicesByField: ReadonlyMap<number, readonly number[]>
  placedFields: ReadonlySet<number>
}

export function buildPivotFieldListIndex(model: Pick<PivotMenuModel, 'fields' | 'data'>): PivotFieldListIndex {
  const fieldsByIndex = new Map<number, PivotMenuField>()
  const dataIndicesByField = new Map<number, number[]>()
  const placedFields = new Set<number>()
  for (const field of model.fields) {
    fieldsByIndex.set(field.index, field)
    if (field.axis !== 'none') placedFields.add(field.index)
  }
  for (const [index, entry] of model.data.entries()) {
    const indices = dataIndicesByField.get(entry.field) ?? []
    indices.push(index)
    dataIndicesByField.set(entry.field, indices)
    placedFields.add(entry.field)
  }
  return { fieldsByIndex, dataIndicesByField, placedFields }
}
