import axios from 'axios';
import { runClientRecovery } from './clientRecoverySolver';

/**
 * File: api.js
 * Author: Antigravity AI
 * Purpose: Axios API service declarations for communicating with FastAPI backend.
 *          Includes robust URL resolution for Render, extended timeouts for cold starts,
 *          and automatic fallback to client-side graph algorithms when the backend is offline.
 */

export const getApiBaseUrl = () => {
  // 1. Explicit user override in settings
  if (typeof window !== 'undefined') {
    const custom = localStorage.getItem('netshield_backend_url');
    if (custom && custom.trim()) {
      let trimmed = custom.trim().replace(/\/+$/, '');
      if (!trimmed.startsWith('http://') && !trimmed.startsWith('https://')) {
        trimmed = `https://${trimmed}`;
      }
      return trimmed;
    }
  }

  // 2. Build-time Vite environment variable
  let envUrl = (import.meta.env.VITE_API_BASE_URL || '').trim().replace(/\/+$/, '');
  if (envUrl) {
    if (!envUrl.startsWith('http://') && !envUrl.startsWith('https://')) {
      envUrl = `https://${envUrl}`;
    }
    return envUrl;
  }

  // 3. Render cloud automatic domain pairing:
  // If hosted on `https://<app>-frontend*.onrender.com`, auto-route to `https://<app>-backend*.onrender.com`
  if (typeof window !== 'undefined' && window.location.hostname.includes('.onrender.com')) {
    if (window.location.hostname.includes('-frontend')) {
      const backendHost = window.location.hostname.replace('-frontend', '-backend');
      return `https://${backendHost}`;
    }
    return 'https://netshield-backend.onrender.com';
  }

  // 4. Fallback for public HTTPS deployments (e.g. Vercel, Netlify)
  if (typeof window !== 'undefined' && window.location.protocol === 'https:') {
    return 'https://netshield-backend.onrender.com';
  }

  // 5. Localhost development
  return 'http://127.0.0.1:8000';
};

const client = axios.create({
  baseURL: getApiBaseUrl(),
  timeout: 60000, // 60 seconds to accommodate Render free tier cold-start wakeups
  headers: {
    'Content-Type': 'application/json',
  },
});

// Update baseURL dynamically per request
client.interceptors.request.use((config) => {
  const currentBase = getApiBaseUrl();
  if (currentBase) {
    config.baseURL = currentBase;
  }
  return config;
});

// Validate that responses are JSON, not HTML from SPA rewrites
client.interceptors.response.use(
  (response) => {
    if (
      typeof response.data === 'string' &&
      (response.data.trim().startsWith('<!DOCTYPE') || response.data.trim().startsWith('<html'))
    ) {
      const err = new Error('Received HTML response instead of JSON. The backend API URL may be misconfigured or pointing to the frontend static site.');
      err.isHtmlResponse = true;
      return Promise.reject(err);
    }
    return response;
  },
  (error) => Promise.reject(error)
);

export const templateService = {
  getTemplates: async () => {
    try {
      const response = await client.get('/templates');
      return response.data;
    } catch (err) {
      console.warn('[templateService] Backend offline, returning standard templates', err.message);
      return {
        templates: [
          { id: 'office_network', name: 'Office Network', description: 'Office environment with workstations, routers, and central server' },
          { id: 'enterprise_tree', name: 'Enterprise Tree', description: 'Hierarchical campus topology with core & distribution switches' },
          { id: 'mesh_backbone', name: 'Mesh Backbone', description: 'Redundant mesh core topology with multi-path failovers' }
        ]
      };
    }
  },
  loadTemplate: async (templateId) => {
    const response = await client.post('/template/load', { template: templateId });
    return response.data;
  },
};

export const algorithmService = {
  getAlgorithms: async () => {
    const response = await client.get('/algorithms');
    return response.data;
  },
  simulateAttack: async (algoId, nodes, edges, startNodes) => {
    const formattedNodes = nodes.map(node => ({
      id: node.id,
      label: node.data?.label || node.id,
      type: node.type || 'PC',
      status: node.data?.status || 'healthy',
      position: {
        x: node.position?.x ?? 0,
        y: node.position?.y ?? 0
      }
    }));

    const formattedEdges = edges.map(edge => ({
      source: edge.source,
      target: edge.target,
      weight: parseFloat(edge.data?.weight ?? 1.0),
      latency: parseFloat(edge.data?.latency ?? 10.0),
      bandwidth: parseFloat(edge.data?.bandwidth ?? 100.0)
    }));

    const response = await client.post('/simulate', {
      algorithm: algoId,
      graph: {
        nodes: formattedNodes,
        edges: formattedEdges
      },
      startNodes
    });
    return response.data;
  },
  recoverNetwork: async (algoId, nodes, edges, options = {}) => {
    const formattedNodes = nodes.map(node => ({
      id: node.id,
      label: node.data?.label || node.id,
      type: node.type || 'PC',
      status: node.data?.status || 'healthy',
      position: {
        x: node.position?.x ?? 0,
        y: node.position?.y ?? 0
      }
    }));

    const formattedEdges = edges.map(edge => ({
      source: edge.source,
      target: edge.target,
      weight: parseFloat(edge.data?.weight ?? 1.0),
      latency: parseFloat(edge.data?.latency ?? 10.0),
      bandwidth: parseFloat(edge.data?.bandwidth ?? 100.0)
    }));

    // 1. Try FastAPI cloud/local backend first
    try {
      const response = await client.post('/recover', {
        algorithm: algoId,
        graph: {
          nodes: formattedNodes,
          edges: formattedEdges
        },
        options
      });

      if (response && response.data && response.data.success) {
        return response.data;
      }
      throw new Error(response?.data?.result?.error || 'Backend calculation returned non-success');
    } catch (err) {
      console.warn(`[RecoveryPlanner] Backend unavailable or failed (${err.message}). Using local algorithmic recovery engine...`);
      // 2. Seamlessly execute client-side algorithmic recovery solver
      return runClientRecovery(algoId, formattedNodes, formattedEdges, options);
    }
  },
  simulateSort: async (algoId, array) => {
    const response = await client.post('/simulate/sort', {
      algorithm: algoId,
      options: { array }
    });
    return response.data;
  },
  simulateDP: async (dimensions) => {
    const response = await client.post('/simulate/dp', {
      options: { dimensions }
    });
    return response.data;
  },
  simulateStrassen: async (A, B) => {
    const response = await client.post('/simulate/strassen', {
      options: { A, B }
    });
    return response.data;
  },
  simulateNQueens: async (N) => {
    const response = await client.post('/simulate/nqueens', {
      options: { N }
    });
    return response.data;
  },
};

export const validationService = {
  validateGraph: async (nodes, edges) => {
    const formattedNodes = nodes.map(node => ({
      id: node.id,
      label: node.data?.label || node.id,
      type: node.type || 'PC',
      status: node.data?.status || 'healthy',
      position: {
        x: node.position?.x ?? 0,
        y: node.position?.y ?? 0
      }
    }));

    const formattedEdges = edges.map(edge => ({
      source: edge.source,
      target: edge.target,
      weight: parseFloat(edge.data?.weight ?? 1.0),
      latency: parseFloat(edge.data?.latency ?? 10.0),
      bandwidth: parseFloat(edge.data?.bandwidth ?? 100.0)
    }));

    try {
      const response = await client.post('/validate', {
        nodes: formattedNodes,
        edges: formattedEdges
      });
      return response.data;
    } catch (err) {
      // Local validation fallback
      const errors = [];
      const warnings = [];
      if (nodes.length === 0) errors.push('Network must contain at least one device.');
      if (edges.length === 0) warnings.push('Network has no connecting links.');
      return { valid: errors.length === 0, errors, warnings };
    }
  },
};

export const projectService = {
  saveProject: async (projectState) => {
    try {
      const response = await client.post('/save', projectState);
      return response.data;
    } catch (err) {
      // Offline fallback: save to localStorage
      localStorage.setItem('netshield_saved_project', JSON.stringify(projectState));
      return { success: true, message: 'Saved locally in browser storage.' };
    }
  },
  loadProject: async () => {
    try {
      const response = await client.post('/load');
      return response.data;
    } catch (err) {
      const saved = localStorage.getItem('netshield_saved_project');
      if (saved) return JSON.parse(saved);
      throw err;
    }
  },
};

export default client;
