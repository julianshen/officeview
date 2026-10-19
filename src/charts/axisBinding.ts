/**
 * Axis-binding validation (spec 5.1). A group's binding must match its
 * family's kind and name existing axes of the right kind; a cartesian
 * binding has exactly one of `category` or `x`. The parser admits a model
 * only when every binding is valid; otherwise it is unsupported (never
 * drawn against a guessed axis).
 */
import type { ChartAxis, ChartAxisBinding, ChartFamily, ChartModel, ChartResult } from './model'

const BINDING_KIND: Readonly<Record<ChartFamily, ChartAxisBinding['kind']>> = {
  bar: 'cartesian', line: 'cartesian', area: 'cartesian', stock: 'cartesian',
  waterfall: 'cartesian', histogram: 'cartesian', pareto: 'cartesian', boxWhisker: 'cartesian', funnel: 'cartesian',
  pie: 'none', doughnut: 'none', ofPie: 'none', treemap: 'none', sunburst: 'none',
  scatter: 'xy', bubble: 'xy',
  radar: 'radial',
  contour: 'surface',
}

const CATEGORY_KINDS: readonly ChartAxis['kind'][] = ['category', 'date']

/** The reason a binding is invalid for the family over these axes, or undefined when it is valid. */
export function validateAxisBinding(family: ChartFamily, binding: ChartAxisBinding, axes: readonly ChartAxis[]): string | undefined {
  const expected = BINDING_KIND[family]
  if (!expected) return `unknown chart family ${String(family)}`
  if (binding.kind !== expected) return `${family} needs a ${expected} axis binding, not ${binding.kind}`
  const byId = new Map(axes.map(axis => [axis.id, axis]))
  const check = (role: string, id: string | undefined, kinds: readonly ChartAxis['kind'][]): string | undefined => {
    const axis = id === undefined ? undefined : byId.get(id)
    if (!axis) return `${family} ${role} axis ${id ?? '(none)'} does not exist`
    return kinds.includes(axis.kind) ? undefined : `${family} ${role} axis ${id} is a ${axis.kind} axis`
  }
  const distinct = (...ids: string[]) => new Set(ids).size === ids.length ? undefined : `${family} binds one axis to two roles`
  switch (binding.kind) {
    case 'none': return undefined
    case 'cartesian':
      if ((binding.category === undefined) === (binding.x === undefined)) return `${family} cartesian binding needs exactly one of category or x`
      return (binding.category !== undefined ? check('category', binding.category, CATEGORY_KINDS) : check('x', binding.x, ['value']))
        ?? check('value', binding.value, ['value']) ?? distinct((binding.category ?? binding.x)!, binding.value)
    case 'xy': return check('x', binding.x, ['value']) ?? check('y', binding.y, ['value']) ?? distinct(binding.x, binding.y)
    case 'radial': return check('value', binding.value, ['value'])
    case 'surface':
      return check('category', binding.category, CATEGORY_KINDS) ?? check('series', binding.series, ['series']) ?? check('value', binding.value, ['value'])
        ?? distinct(binding.category, binding.series, binding.value)
  }
}

/** Parse-time admission: `native` when every axis id is unique and every group binding is valid. */
export function admitChartModel(model: ChartModel): ChartResult {
  const ids = new Set<string>()
  for (const axis of model.axes) {
    if (ids.has(axis.id)) return { kind: 'unsupported', reason: `duplicate chart axis id ${axis.id}` }
    ids.add(axis.id)
  }
  for (const group of model.groups) {
    const reason = validateAxisBinding(group.family, group.binding, model.axes)
    if (reason) return { kind: 'unsupported', reason }
  }
  return { kind: 'native', model }
}
