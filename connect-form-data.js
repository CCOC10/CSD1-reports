// connect-form-data.js
// Simple JS module for storing form data from CSD1 connect.html
// Usage: import or include in CSD1 connect.html

const ConnectFormDataStore = {
  data: {},

  set(key, value) {
    this.data[key] = value;
  },

  get(key) {
    return this.data[key];
  },

  getAll() {
    return { ...this.data };
  },

  clear() {
    this.data = {};
  }
};

// Example: ConnectFormDataStore.set('adminNoteTag', 'urgent_special');
// Example: ConnectFormDataStore.get('adminNoteTag');
// Example: ConnectFormDataStore.getAll();
// Example: ConnectFormDataStore.clear();

// Export for module usage
if (typeof module !== 'undefined') {
  module.exports = ConnectFormDataStore;
}
