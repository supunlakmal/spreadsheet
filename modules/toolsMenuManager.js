/**
 * Tools Menu Manager
 * Handles dynamic rendering of tools menu from JSON configuration
 */

export class ToolsMenuManager {
  static config = null;
  static searchInput = null;
  static gridContainer = null;
  static emptyState = null;

  /**
   * Initialize the tools menu manager
   */
  static async init() {
    try {
      // Load configuration
      const response = await fetch('./tools-config.json');
      if (!response.ok) {
        throw new Error('Failed to load tools configuration');
      }
      this.config = await response.json();

      // Get DOM elements
      this.searchInput = document.getElementById('tools-search-input');
      this.gridContainer = document.getElementById('tools-dynamic-grid');
      this.emptyState = document.getElementById('tools-empty-state');

      if (!this.gridContainer) {
        console.error('Tools grid container not found');
        return;
      }

      // Set up search functionality
      if (this.searchInput) {
        this.searchInput.addEventListener('input', (e) => {
          this.filterTools(e.target.value);
        });
      }

      // Initial render
      this.render();
    } catch (error) {
      console.error('Error initializing ToolsMenuManager:', error);
      // Fallback: show error message in grid
      if (this.gridContainer) {
        this.gridContainer.innerHTML = `
          <div style="padding: 20px; text-align: center; color: var(--text-muted);">
            <i class="fa-solid fa-exclamation-triangle"></i>
            <p>Failed to load tools configuration</p>
          </div>
        `;
      }
    }
  }

  /**
   * Render all tools based on current config and filter
   */
  static render(filterQuery = '') {
    if (!this.config || !this.gridContainer) return;

    const { categories, tools } = this.config;
    const query = filterQuery.toLowerCase().trim();

    // Filter tools based on search query
    const filteredTools = tools.filter(tool => {
      if (!query) return true;
      return (
        tool.name.toLowerCase().includes(query) ||
        tool.description.toLowerCase().includes(query) ||
        tool.category.toLowerCase().includes(query)
      );
    });

    // Show/hide empty state
    if (this.emptyState) {
      if (filteredTools.length === 0) {
        this.emptyState.classList.remove('hidden');
        this.gridContainer.innerHTML = '';
        return;
      } else {
        this.emptyState.classList.add('hidden');
      }
    }

    // Sort categories by order
    const sortedCategories = [...categories].sort((a, b) => a.order - b.order);

    // Build HTML
    let html = '';

    sortedCategories.forEach(category => {
      // Get tools for this category
      const categoryTools = filteredTools
        .filter(tool => tool.category === category.id)
        .sort((a, b) => a.order - b.order);

      if (categoryTools.length === 0) return;

      // Category section
      html += `
        <div class="tools-section">
          <h4><i class="fa-solid ${category.icon}"></i> ${category.label}</h4>
          <div class="tools-list">
      `;

      // Tools in this category
      categoryTools.forEach(tool => {
        html += `
          <button 
            id="${tool.id}" 
            type="button" 
            class="tool-item" 
            title="${tool.description}"
            data-tool-id="${tool.id}"
          >
            <i class="fa-solid ${tool.icon}"></i>
            ${tool.name}
          </button>
        `;
      });

      html += `
          </div>
        </div>
      `;
    });

    this.gridContainer.innerHTML = html;
  }

  /**
   * Filter tools by search query
   */
  static filterTools(query) {
    this.render(query);
  }

  /**
   * Add a new tool programmatically
   */
  static addTool(tool) {
    if (!this.config) return;
    this.config.tools.push(tool);
    this.render();
  }

  /**
   * Remove a tool by ID
   */
  static removeTool(toolId) {
    if (!this.config) return;
    this.config.tools = this.config.tools.filter(t => t.id !== toolId);
    this.render();
  }

  /**
   * Get tool configuration by ID
   */
  static getTool(toolId) {
    if (!this.config) return null;
    return this.config.tools.find(t => t.id === toolId);
  }

  /**
   * Update tool configuration
   */
  static updateTool(toolId, updates) {
    if (!this.config) return;
    const tool = this.config.tools.find(t => t.id === toolId);
    if (tool) {
      Object.assign(tool, updates);
      this.render();
    }
  }

  /**
   * Get all categories
   */
  static getCategories() {
    return this.config ? this.config.categories : [];
  }

  /**
   * Get all tools
   */
  static getTools() {
    return this.config ? this.config.tools : [];
  }
}
