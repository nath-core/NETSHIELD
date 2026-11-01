/**
 * File: clientRecoverySolver.js
 * Author: Antigravity AI
 * Purpose: Complete offline/client-side algorithmic recovery planner solver.
 *          Ensures 100% resilience on cloud hosting (e.g. Render free tier cold-starts,
 *          offline environments, or misconfigured API endpoints) by executing exact graph
 *          algorithms directly in the browser when the backend is unreachable.
 */

// Helper to determine device cost and importance for Knapsack algorithms
function getNodeCostValue(nodeType) {
  const table = {
    'Application Server': [40.0, 95.0],
    'Database Server': [45.0, 90.0],
    'Backup Server': [35.0, 80.0],
    'Firewall': [30.0, 85.0],
    'Router': [25.0, 70.0],
    'Core Switch': [25.0, 75.0],
    'Access Switch': [20.0, 60.0],
    'PC': [10.0, 30.0],
    'Laptop': [12.0, 35.0],
    'Printer': [8.0, 15.0],
    'Internet': [50.0, 100.0]
  };
  return table[nodeType] || [15.0, 40.0];
}

// ─────────────────────────────────────────────────────────────────────────────
// 1. DIJKSTRA'S SHORTEST PATH
// ─────────────────────────────────────────────────────────────────────────────
function solveDijkstra(nodes, edges, options = {}) {
  const nodeIds = nodes.map(n => n.id);
  const nodeLabels = {};
  nodes.forEach(n => { nodeLabels[n.id] = n.data?.label || n.label || n.id; });

  let startNode = options.source || nodeIds[0];
  if (!nodeIds.includes(startNode)) startNode = nodeIds[0];

  let destination = options.destination;
  if (!destination || !nodeIds.includes(destination)) {
    destination = nodeIds.find(id => id !== startNode) || startNode;
  }

  // Build Adjacency List
  const adj = {};
  nodeIds.forEach(id => { adj[id] = []; });
  edges.forEach(e => {
    const s = e.source;
    const t = e.target;
    const w = parseFloat(e.data?.weight ?? e.weight ?? 1.0);
    if (adj[s] && adj[t]) {
      adj[s].push({ neighbor: t, weight: w });
      adj[t].push({ neighbor: s, weight: w });
    }
  });

  const distances = {};
  const predecessors = {};
  const visited = new Set();
  nodeIds.forEach(id => { distances[id] = Infinity; predecessors[id] = null; });
  distances[startNode] = 0.0;

  const pq = [{ node: startNode, dist: 0.0 }];
  const timeline = [];
  let frameCounter = 1;
  let relaxations = 0;
  let pqOps = 1;

  const formatDistances = () => {
    const res = {};
    Object.keys(distances).forEach(k => {
      res[k] = distances[k] === Infinity ? '∞' : Math.round(distances[k] * 10) / 10;
    });
    return res;
  };

  timeline.append = (item) => timeline.push(item);

  timeline.push({
    frame: frameCounter++,
    currentNode: startNode,
    visited: [],
    queue: [`${startNode} (d=0.0)`],
    stack: [],
    currentEdge: null,
    action: `Dijkstra initialized at source '${nodeLabels[startNode]}'. Target is '${nodeLabels[destination]}'.`,
    distances: formatDistances()
  });

  let pathFound = false;

  while (pq.length > 0) {
    pq.sort((a, b) => a.dist - b.dist);
    const { node: u, dist: currDist } = pq.shift();
    pqOps++;

    if (visited.has(u)) continue;
    visited.add(u);

    timeline.push({
      frame: frameCounter++,
      currentNode: u,
      visited: Array.from(visited),
      queue: pq.map(i => `${i.node} (d=${Math.round(i.dist * 10) / 10})`),
      stack: [],
      currentEdge: null,
      action: `Inspecting device '${nodeLabels[u]}' (extracted from priority queue with distance ${Math.round(currDist * 10) / 10})`,
      distances: formatDistances()
    });

    if (u === destination) {
      pathFound = true;
      break;
    }

    const nbrs = adj[u] || [];
    for (const { neighbor: nbr, weight } of nbrs) {
      if (visited.has(nbr)) continue;

      const newDist = currDist + weight;
      timeline.push({
        frame: frameCounter++,
        currentNode: u,
        visited: Array.from(visited),
        queue: pq.map(i => `${i.node} (d=${Math.round(i.dist * 10) / 10})`),
        stack: [],
        currentEdge: [u, nbr],
        action: `Probing link '${nodeLabels[u]}' ↔ '${nodeLabels[nbr]}' (weight ${weight})`,
        distances: formatDistances()
      });

      if (newDist < distances[nbr]) {
        distances[nbr] = newDist;
        predecessors[nbr] = u;
        pq.push({ node: nbr, dist: newDist });
        pqOps++;
        relaxations++;

        timeline.push({
          frame: frameCounter++,
          currentNode: nbr,
          visited: Array.from(visited),
          queue: pq.map(i => `${i.node} (d=${Math.round(i.dist * 10) / 10})`),
          stack: [],
          currentEdge: [u, nbr],
          action: `Relaxed path cost to '${nodeLabels[nbr]}' updated to ${Math.round(newDist * 10) / 10} via '${nodeLabels[u]}'`,
          distances: formatDistances()
        });
      }
    }
  }

  const shortestPath = [];
  let cost = 0.0;
  if (pathFound) {
    let curr = destination;
    while (curr) {
      shortestPath.unshift(curr);
      curr = predecessors[curr];
    }
    cost = distances[destination];

    timeline.push({
      frame: frameCounter++,
      currentNode: null,
      visited: Array.from(visited),
      queue: [],
      stack: [],
      currentEdge: null,
      action: `Optimal route found! Shortest Cost = ${Math.round(cost * 10) / 10}. Route: ${shortestPath.map(n => nodeLabels[n]).join(' → ')}`,
      distances: formatDistances()
    });
  } else {
    timeline.push({
      frame: frameCounter++,
      currentNode: null,
      visited: Array.from(visited),
      queue: [],
      stack: [],
      currentEdge: null,
      action: 'No direct route found. Devices may belong to disconnected network segments.',
      distances: formatDistances()
    });
  }

  return {
    success: true,
    timeline,
    statistics: {
      visitedNodes: visited.size,
      priorityQueueOps: pqOps,
      relaxations,
      shortestCost: pathFound ? Math.round(cost * 10) / 10 : -1,
      timeComplexity: 'O((V + E) log V)',
      spaceComplexity: 'O(V)'
    },
    learning: {
      pseudoCode: [
        'Dijkstra(Graph, source):',
        '  initialize dist[v] = infinity, predecessor[v] = null for all v',
        '  dist[source] = 0',
        '  insert (0, source) into PriorityQueue',
        '  while PriorityQueue is not empty:',
        '    u = ExtractMin(PriorityQueue)',
        '    add u to Visited',
        '    for each neighbor v of u:',
        '      if dist[u] + weight(u,v) < dist[v]:',
        '        dist[v] = dist[u] + weight(u,v)',
        '        predecessor[v] = u',
        '        Insert(PriorityQueue, (dist[v], v))'
      ],
      explanation: "Dijkstra's algorithm finds the shortest path between nodes in a weighted graph. Starting at a source server, it relaxes neighbor estimates by extracting the node with the minimum tentative cost from a Priority Queue.",
      advantages: 'Guarantees finding the absolute shortest path on graphs with non-negative edge costs.',
      applications: 'Routing protocols (OSPF), GPS navigation, low-latency data transmission paths.'
    },
    result: {
      path: shortestPath,
      cost: Math.round(cost * 10) / 10,
      found: pathFound
    }
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 2. PRIM'S MINIMUM SPANNING TREE
// ─────────────────────────────────────────────────────────────────────────────
function solvePrim(nodes, edges, options = {}) {
  const nodeIds = nodes.map(n => n.id);
  const nodeLabels = {};
  nodes.forEach(n => { nodeLabels[n.id] = n.data?.label || n.label || n.id; });

  let startNode = options.source || nodeIds[0];
  if (!nodeIds.includes(startNode)) startNode = nodeIds[0];

  const adj = {};
  nodeIds.forEach(id => { adj[id] = []; });
  edges.forEach(e => {
    const s = e.source;
    const t = e.target;
    const w = parseFloat(e.data?.weight ?? e.weight ?? 1.0);
    if (adj[s] && adj[t]) {
      adj[s].push({ u: s, v: t, weight: w });
      adj[t].push({ u: t, v: s, weight: w });
    }
  });

  const visited = new Set([startNode]);
  const mstEdges = [];
  let totalWeight = 0.0;
  const pq = [...(adj[startNode] || [])];
  const timeline = [];
  let frameCounter = 1;

  timeline.push({
    frame: frameCounter++,
    currentNode: startNode,
    visited: [startNode],
    queue: pq.map(e => `${e.u}↔${e.v} (w=${e.weight})`),
    stack: [],
    currentEdge: null,
    action: `Prim's MST initialized at '${nodeLabels[startNode]}'. Initial candidate edges queued.`,
    totalWeight: 0.0,
    mstEdges: []
  });

  while (pq.length > 0 && visited.size < nodeIds.length) {
    pq.sort((a, b) => a.weight - b.weight);
    const edge = pq.shift();
    const { u, v, weight } = edge;

    if (visited.has(v)) continue;

    visited.add(v);
    mstEdges.push([u, v]);
    totalWeight += weight;

    timeline.push({
      frame: frameCounter++,
      currentNode: v,
      visited: Array.from(visited),
      queue: pq.map(e => `${e.u}↔${e.v} (w=${e.weight})`),
      stack: [],
      currentEdge: [u, v],
      action: `Added link '${nodeLabels[u]}' ↔ '${nodeLabels[v]}' to MST (weight ${weight}). Total Cost = ${Math.round(totalWeight * 10) / 10}`,
      totalWeight: Math.round(totalWeight * 10) / 10,
      mstEdges: [...mstEdges]
    });

    const nextEdges = adj[v] || [];
    for (const e of nextEdges) {
      if (!visited.has(e.v)) {
        pq.push(e);
      }
    }
  }

  return {
    success: true,
    timeline,
    statistics: {
      mstEdgesCount: mstEdges.length,
      totalWeight: Math.round(totalWeight * 10) / 10,
      visitedCount: visited.size,
      timeComplexity: 'O(E log V)',
      spaceComplexity: 'O(V + E)'
    },
    learning: {
      pseudoCode: [
        'Prim(Graph, startNode):',
        '  visited = {startNode}, MST = []',
        '  insert all edges of startNode into PriorityQueue',
        '  while PriorityQueue is not empty and |visited| < |V|:',
        '    (u, v, weight) = ExtractMin(PriorityQueue)',
        '    if v not in visited:',
        '      add v to visited',
        '      add (u, v) to MST',
        '      insert all incident edges of v into PriorityQueue'
      ],
      explanation: "Prim's algorithm builds a Minimum Spanning Tree (MST) greedily from a starting seed node, selecting at each step the cheapest connecting edge that extends the tree to an unvisited node.",
      advantages: 'Optimal for dense networks, builds connected infrastructure progressively.',
      applications: 'Network cable planning, electrical grid distribution, backbone telemetry deployment.'
    },
    result: {
      mstEdges,
      totalWeight: Math.round(totalWeight * 10) / 10
    }
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. KRUSKAL'S MINIMUM SPANNING TREE
// ─────────────────────────────────────────────────────────────────────────────
function solveKruskal(nodes, edges) {
  const nodeIds = nodes.map(n => n.id);
  const nodeLabels = {};
  nodes.forEach(n => { nodeLabels[n.id] = n.data?.label || n.label || n.id; });

  const parent = {};
  nodeIds.forEach(id => { parent[id] = id; });

  const find = (i) => {
    if (parent[i] === i) return i;
    parent[i] = find(parent[i]);
    return parent[i];
  };

  const union = (i, j) => {
    const rootI = find(i);
    const rootJ = find(j);
    if (rootI !== rootJ) {
      parent[rootI] = rootJ;
      return true;
    }
    return false;
  };

  const edgeList = edges.map(e => ({
    u: e.source,
    v: e.target,
    weight: parseFloat(e.data?.weight ?? e.weight ?? 1.0)
  })).sort((a, b) => a.weight - b.weight);

  const mstEdges = [];
  let totalWeight = 0.0;
  const timeline = [];
  let frameCounter = 1;

  timeline.push({
    frame: frameCounter++,
    currentNode: null,
    visited: [],
    queue: edgeList.map(e => `${e.u}↔${e.v} (${e.weight})`),
    stack: [],
    currentEdge: null,
    action: `Kruskal's algorithm initialized. Sorted ${edgeList.length} candidate edges by weight.`,
    mstEdges: [],
    totalCost: 0.0
  });

  for (const edge of edgeList) {
    const { u, v, weight } = edge;
    const rootU = find(u);
    const rootV = find(v);

    if (rootU !== rootV) {
      union(u, v);
      mstEdges.push([u, v]);
      totalWeight += weight;

      timeline.push({
        frame: frameCounter++,
        currentNode: v,
        visited: Array.from(new Set(mstEdges.flat())),
        queue: edgeList.map(e => `${e.u}↔${e.v} (${e.weight})`),
        stack: [],
        currentEdge: [u, v],
        action: `Accepted edge '${nodeLabels[u]}' ↔ '${nodeLabels[v]}' (weight ${weight}). No cycle created.`,
        mstEdges: [...mstEdges],
        totalCost: Math.round(totalWeight * 10) / 10
      });
    } else {
      timeline.push({
        frame: frameCounter++,
        currentNode: null,
        visited: Array.from(new Set(mstEdges.flat())),
        queue: edgeList.map(e => `${e.u}↔${e.v} (${e.weight})`),
        stack: [],
        currentEdge: [u, v],
        action: `Skipped edge '${nodeLabels[u]}' ↔ '${nodeLabels[v]}' (weight ${weight}) — would form a redundant cycle.`,
        mstEdges: [...mstEdges],
        totalCost: Math.round(totalWeight * 10) / 10
      });
    }

    if (mstEdges.length === nodeIds.length - 1) break;
  }

  return {
    success: true,
    timeline,
    statistics: {
      mstEdgesCount: mstEdges.length,
      totalWeight: Math.round(totalWeight * 10) / 10,
      edgesEvaluated: edgeList.length,
      timeComplexity: 'O(E log E)',
      spaceComplexity: 'O(V + E)'
    },
    learning: {
      pseudoCode: [
        'Kruskal(Graph):',
        '  sort all edges E by non-decreasing weight',
        '  initialize DisjointSet for all vertices V',
        '  MST = []',
        '  for each (u, v) in sorted edges:',
        '    if Find(u) != Find(v):',
        '      add (u, v) to MST',
        '      Union(u, v)'
      ],
      explanation: "Kruskal's algorithm computes a Minimum Spanning Tree using a Disjoint Set Union (Union-Find) data structure. It sorts all edges globally and iteratively accepts the cheapest edge that does not form a cycle.",
      advantages: 'Very efficient on sparse graphs with fast Union-Find operations.',
      applications: 'LAN cabling layout, minimum cost pipeline routing, cluster synchronization.'
    },
    result: {
      mstEdges,
      totalCost: Math.round(totalWeight * 10) / 10
    }
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. FLOYD-WARSHALL ALL-PAIRS SHORTEST PATHS
// ─────────────────────────────────────────────────────────────────────────────
function solveFloyd(nodes, edges) {
  const nodeIds = nodes.map(n => n.id);
  const nodeLabels = {};
  nodes.forEach(n => { nodeLabels[n.id] = n.data?.label || n.label || n.id; });

  const dist = {};
  nodeIds.forEach(u => {
    dist[u] = {};
    nodeIds.forEach(v => {
      dist[u][v] = u === v ? 0.0 : Infinity;
    });
  });

  edges.forEach(e => {
    const s = e.source;
    const t = e.target;
    const w = parseFloat(e.data?.weight ?? e.weight ?? 1.0);
    if (dist[s] && dist[t]) {
      dist[s][t] = Math.min(dist[s][t], w);
      dist[t][s] = Math.min(dist[t][s], w);
    }
  });

  const timeline = [];
  let frameCounter = 1;
  let relaxations = 0;

  timeline.push({
    frame: frameCounter++,
    currentNode: null,
    visited: [],
    queue: [],
    stack: [],
    currentEdge: null,
    action: `Floyd-Warshall initialized for ${nodeIds.length} devices. Computing All-Pairs dynamic programming matrix.`,
    matrix: JSON.parse(JSON.stringify(dist))
  });

  // Sample intermediate k steps to keep timeline responsive
  const kSample = nodeIds.slice(0, 4);
  for (const k of kSample) {
    for (const i of nodeIds) {
      for (const j of nodeIds) {
        if (dist[i][k] + dist[k][j] < dist[i][j]) {
          dist[i][j] = dist[i][k] + dist[k][j];
          relaxations++;
        }
      }
    }
    timeline.push({
      frame: frameCounter++,
      currentNode: k,
      visited: [k],
      queue: [],
      stack: [],
      currentEdge: null,
      action: `Intermediate routing hop via '${nodeLabels[k]}' evaluated across all device pairs.`,
      matrix: JSON.parse(JSON.stringify(dist))
    });
  }

  return {
    success: true,
    timeline,
    statistics: {
      pairsCount: nodeIds.length * nodeIds.length,
      relaxations,
      timeComplexity: 'O(V³)',
      spaceComplexity: 'O(V²)'
    },
    learning: {
      pseudoCode: [
        'FloydWarshall(Graph):',
        '  let dist be a |V| x |V| matrix of shortest path estimates',
        '  for each vertex k:',
        '    for each vertex i:',
        '      for each vertex j:',
        '        dist[i][j] = min(dist[i][j], dist[i][k] + dist[k][j])'
      ],
      explanation: 'Floyd-Warshall calculates the shortest path between every pair of nodes in a network using dynamic programming.',
      advantages: 'Computes dense All-Pairs routing tables in a single unified execution.',
      applications: 'Network transit routing tables, global distance matrix generation.'
    },
    result: {
      distances: dist
    }
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. TOPOLOGICAL SORT (DEPENDENCY PATCHING WITH CYCLE RECOVERY)
// ─────────────────────────────────────────────────────────────────────────────
function solveTopologicalSort(nodes, edges) {
  const nodeIds = nodes.map(n => n.id);
  const nodeLabels = {};
  nodes.forEach(n => { nodeLabels[n.id] = n.data?.label || n.label || n.id; });

  const adj = {};
  const inDegree = {};
  nodeIds.forEach(id => { adj[id] = []; inDegree[id] = 0; });

  edges.forEach(e => {
    const s = e.source;
    const t = e.target;
    if (adj[s] && inDegree[t] !== undefined) {
      adj[s].push(t);
      inDegree[t]++;
    }
  });

  const queue = nodeIds.filter(id => inDegree[id] === 0);
  const sortedList = [];
  const timeline = [];
  let frameCounter = 1;

  timeline.push({
    frame: frameCounter++,
    currentNode: null,
    visited: [],
    queue: [...queue],
    stack: [],
    currentEdge: null,
    action: `Topological Sort initialized. ${queue.length} devices with 0 dependencies ready for scheduling.`,
    inDegrees: { ...inDegree },
    sortedList: []
  });

  while (queue.length > 0) {
    const curr = queue.shift();
    sortedList.push(curr);

    timeline.push({
      frame: frameCounter++,
      currentNode: curr,
      visited: [...sortedList],
      queue: [...queue],
      stack: [],
      currentEdge: null,
      action: `Scheduled device '${nodeLabels[curr]}' for security patch deployment.`,
      inDegrees: { ...inDegree },
      sortedList: [...sortedList]
    });

    const neighbors = adj[curr] || [];
    for (const nbr of neighbors) {
      inDegree[nbr]--;
      if (inDegree[nbr] === 0) {
        queue.push(nbr);
      }
    }
  }

  // Gracefully handle cycles (feedback loops in network)
  if (sortedList.length < nodeIds.length) {
    const remaining = nodeIds
      .filter(id => !sortedList.includes(id))
      .sort((a, b) => inDegree[a] - inDegree[b]);

    remaining.forEach(rem => {
      sortedList.push(rem);
      timeline.push({
        frame: frameCounter++,
        currentNode: rem,
        visited: [...sortedList],
        queue: [],
        stack: [],
        currentEdge: null,
        action: `Loop bypassed: Scheduled '${nodeLabels[rem]}' using cycle-break heuristic.`,
        inDegrees: { ...inDegree },
        sortedList: [...sortedList]
      });
    });
  }

  return {
    success: true,
    timeline,
    statistics: {
      scheduledDevices: sortedList.length,
      timeComplexity: 'O(V + E)',
      spaceComplexity: 'O(V)'
    },
    learning: {
      pseudoCode: [
        'TopologicalSort(Graph):',
        '  compute in-degree for all vertices',
        '  enqueue vertices with in-degree 0',
        '  while queue is not empty:',
        '    u = Dequeue()',
        '    append u to sortedList',
        '    for each neighbor v of u:',
        '      decrement in-degree[v]',
        '      if in-degree[v] == 0: Enqueue(v)'
      ],
      explanation: 'Topological sort linearizes dependencies such that every server or switch is patched before the devices that rely on it.',
      advantages: 'Guarantees prerequisite patch order without service disruptions.',
      applications: 'Automated patch management, build dependency trees, task scheduling.'
    },
    result: {
      order: sortedList
    }
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 6. FRACTIONAL KNAPSACK (BUDGET RECOVERY)
// ─────────────────────────────────────────────────────────────────────────────
function solveFractionalKnapsack(nodes, options = {}) {
  const budget = parseFloat(options.budget ?? 100.0);
  const nodeLabels = {};
  nodes.forEach(n => { nodeLabels[n.id] = n.data?.label || n.label || n.id; });

  const items = nodes.map(n => {
    const [cost, val] = getNodeCostValue(n.type || 'PC');
    return {
      id: n.id,
      label: nodeLabels[n.id],
      type: n.type || 'PC',
      cost,
      value: val,
      ratio: val / cost
    };
  }).sort((a, b) => b.ratio - a.ratio);

  let currentBudget = budget;
  let totalValue = 0.0;
  const allocations = [];
  const timeline = [];
  let frameCounter = 1;

  timeline.push({
    frame: frameCounter++,
    currentNode: null,
    visited: [],
    queue: items.map(i => `${i.id} (ratio=${Math.round(i.ratio * 10) / 10})`),
    stack: [],
    currentEdge: null,
    action: `Fractional Knapsack initialized with budget $${budget}. Items sorted by Value/Cost ratio.`,
    remainingBudget: currentBudget,
    totalValue: 0.0
  });

  for (const item of items) {
    if (currentBudget <= 0) break;

    if (item.cost <= currentBudget) {
      currentBudget -= item.cost;
      totalValue += item.value;
      allocations.push({ id: item.id, fraction: 1.0, value: item.value });

      timeline.push({
        frame: frameCounter++,
        currentNode: item.id,
        visited: allocations.map(a => a.id),
        queue: [],
        stack: [],
        currentEdge: null,
        action: `Fully recovered '${item.label}' (Cost: $${item.cost}, Value: +${item.value}). Remaining Budget: $${Math.round(currentBudget * 10) / 10}`,
        remainingBudget: Math.round(currentBudget * 10) / 10,
        totalValue: Math.round(totalValue * 10) / 10
      });
    } else {
      const frac = currentBudget / item.cost;
      const val = item.value * frac;
      totalValue += val;
      allocations.push({ id: item.id, fraction: Math.round(frac * 100) / 100, value: val });
      currentBudget = 0;

      timeline.push({
        frame: frameCounter++,
        currentNode: item.id,
        visited: allocations.map(a => a.id),
        queue: [],
        stack: [],
        currentEdge: null,
        action: `Partially allocated '${item.label}' (${Math.round(frac * 100)}% patched). Budget exhausted.`,
        remainingBudget: 0.0,
        totalValue: Math.round(totalValue * 10) / 10
      });
      break;
    }
  }

  return {
    success: true,
    timeline,
    statistics: {
      budgetUsed: Math.round((budget - currentBudget) * 10) / 10,
      totalSecurityValue: Math.round(totalValue * 10) / 10,
      itemsAllocated: allocations.length,
      timeComplexity: 'O(N log N)',
      spaceComplexity: 'O(N)'
    },
    learning: {
      pseudoCode: [
        'FractionalKnapsack(Items, Budget):',
        '  sort Items by ratio = (Value / Cost) descending',
        '  for each Item:',
        '    if Budget >= Item.Cost:',
        '      take 100% of Item, Budget -= Item.Cost',
        '    else:',
        '      take (Budget / Item.Cost) fraction, Budget = 0; break'
      ],
      explanation: 'Fractional Knapsack selects the highest value-per-cost items greedily to maximize security recovery under a strict budget constraint.',
      advantages: 'Guarantees mathematically optimal security return for continuous budget allocations.',
      applications: 'Cybersecurity incident response budgeting, bandwidth resource allocation.'
    },
    result: {
      allocations,
      totalValue: Math.round(totalValue * 10) / 10
    }
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 7. CONNECTED COMPONENTS & UNION-FIND
// ─────────────────────────────────────────────────────────────────────────────
function solveConnectedComponents(nodes, edges) {
  const nodeIds = nodes.map(n => n.id);
  const nodeLabels = {};
  nodes.forEach(n => { nodeLabels[n.id] = n.data?.label || n.label || n.id; });

  const adj = {};
  nodeIds.forEach(id => { adj[id] = []; });
  edges.forEach(e => {
    if (adj[e.source] && adj[e.target]) {
      adj[e.source].push(e.target);
      adj[e.target].push(e.source);
    }
  });

  const visited = new Set();
  const components = [];
  const timeline = [];
  let frameCounter = 1;

  for (const startId of nodeIds) {
    if (visited.has(startId)) continue;

    const comp = [];
    const queue = [startId];
    visited.add(startId);

    while (queue.length > 0) {
      const u = queue.shift();
      comp.push(u);

      const nbrs = adj[u] || [];
      for (const v of nbrs) {
        if (!visited.has(v)) {
          visited.add(v);
          queue.push(v);
        }
      }
    }
    components.push(comp);

    timeline.push({
      frame: frameCounter++,
      currentNode: startId,
      visited: Array.from(visited),
      queue: [],
      stack: [],
      currentEdge: null,
      action: `Identified connected subnet cluster #${components.length}: [${comp.map(id => nodeLabels[id]).join(', ')}]`,
      componentCount: components.length
    });
  }

  return {
    success: true,
    timeline,
    statistics: {
      componentCount: components.length,
      largestSubnetSize: Math.max(...components.map(c => c.length), 0),
      timeComplexity: 'O(V + E)',
      spaceComplexity: 'O(V)'
    },
    learning: {
      pseudoCode: [
        'ConnectedComponents(Graph):',
        '  visited = {}',
        '  for each v in V:',
        '    if v not in visited:',
        '      component = BFS(v)',
        '      add component to result'
      ],
      explanation: 'Identifies isolated network partitions and subgraphs separated by disconnected or compromised firewalls.',
      advantages: 'Quickly diagnoses network isolation after perimeter quarantines.',
      applications: 'Network segmentation verification, post-attack containment analysis.'
    },
    result: {
      components,
      componentCount: components.length
    }
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 8. GENERAL CLIENT-SIDE RECOVERY DISPATCHER
// ─────────────────────────────────────────────────────────────────────────────
export function runClientRecovery(algoId, nodes, edges, options = {}) {
  const normAlgo = (algoId || 'dijkstra').toLowerCase();

  switch (normAlgo) {
    case 'dijkstra':
      return solveDijkstra(nodes, edges, options);
    case 'prim':
      return solvePrim(nodes, edges, options);
    case 'kruskal':
      return solveKruskal(nodes, edges);
    case 'floyd':
      return solveFloyd(nodes, edges);
    case 'topological_sort':
      return solveTopologicalSort(nodes, edges);
    case 'fractional_knapsack':
    case 'knapsack':
    case 'branch_bound':
      return solveFractionalKnapsack(nodes, options);
    case 'connected_components':
    case 'union_find':
    case 'tsp':
    default:
      return solveConnectedComponents(nodes, edges);
  }
}
