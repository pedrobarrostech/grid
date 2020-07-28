// @ts-nocheck
import { Dag, Node, DependencyMapping } from './../src/graph'

describe('Dag', () => {
  let mapping: DependencyMapping
  let graph: Dag<Node>
  beforeEach(() => {
    mapping = new DependencyMapping()
    graph = new Dag<Node>(node => node.children)
  })
  it('exists', () => {
    expect(Dag).toBeDefined()
  })
  it('navigates the dependency', () => {
    const formula = 'B1=SUM(A1, A2)'
    const cell = 'B1'
    const sheet = 'Sheet1'
    const deps = ['A1', 'A2']
    const cellNode = mapping.get(cell, sheet)
    deps.forEach(dep => {
      const node = mapping.get(dep, sheet)
      node.children.add(cellNode)
    })

    const visited = graph.visit([
      mapping.get(deps[0], sheet)
    ])

    expect(visited.has(cellNode)).toBe(true)
  })

  it('navigates nested dependency', () => {
    const formula = 'B1=SUM(A1, A2)'
    const formula2 = 'B2=SUM(B1, A1)'
    const cell = 'B1'
    const cell2 = 'B2'
    const sheet = 'Sheet1'
    const deps = ['A1', 'A2']
    const deps2 = ['B1', 'A1']
    const cellNode = mapping.get(cell, sheet)
    const cellNode2 = mapping.get(cell2, sheet)

    deps.forEach(dep => {
      const node = mapping.get(dep, sheet)
      node.children.add(cellNode)
    })

    deps2.forEach(dep => {
      const node = mapping.get(dep, sheet)
      node.children.add(cellNode2)
    })

    let visited = graph.visit([
      mapping.get(deps[0], sheet)
    ])

    expect(visited.has(cellNode)).toBe(true)
    expect(visited.has(cellNode2)).toBe(true)

    // A2 => B1 => B2
    visited = graph.visit([
      mapping.get(deps[1], sheet)
    ])

    expect(visited.has(cellNode2)).toBe(true)
  })

  it('allow removal of dependency', () => {
    const formula = 'B1=SUM(A1, A2)'
    const cell = 'B1'
    const sheet = 'Sheet1'
    const deps = ['A1', 'A2']
    const cellNode = mapping.get(cell, sheet)

    deps.forEach(dep => {
      const node = mapping.get(dep, sheet)
      node.children.add(cellNode)
    })

    let visited = graph.visit([
      mapping.get(deps[0], sheet)
    ])

    expect(visited.has(cellNode)).toBe(true)

    // User deletes B1=SUM(A1, A2)
    deps.forEach(dep => {
      const node = mapping.get(dep, sheet)
      node.children.delete(cellNode)
    })

    visited = graph.visit([
      mapping.get(deps[0], sheet)
    ])
    
    expect(visited.has(cellNode)).toBe(false)
  })
})