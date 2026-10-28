/**
 * Spreadsheet format plugin registry.
 * Allows extending the core XLSX reader with additional format groups.
 */

// Global registry of format handlers, writers and sniffers
const registeredParsers = new Map();
const registeredWriters = new Map();
const registeredSniffers = new Set();
const pluginNames = new Set();
const pluginsMap = new Map();

// Mapping of format extension to the plugin that provides it
export const KNOWN_PLUGIN_MAPPINGS = {
    csv: { plugin: 'textPlugin', importPath: 'spreadsheet-utils/plugins' },
    tsv: { plugin: 'textPlugin', importPath: 'spreadsheet-utils/plugins' },
    tab: { plugin: 'textPlugin', importPath: 'spreadsheet-utils/plugins' },
    txt: { plugin: 'textPlugin', importPath: 'spreadsheet-utils/plugins' },
    xls: { plugin: 'xlsPlugin', importPath: 'spreadsheet-utils/plugins' },
    ods: { plugin: 'markupPlugin', importPath: 'spreadsheet-utils/plugins' },
    html: { plugin: 'markupPlugin', importPath: 'spreadsheet-utils/plugins' },
    htm: { plugin: 'markupPlugin', importPath: 'spreadsheet-utils/plugins' },
    xml: { plugin: 'markupPlugin', importPath: 'spreadsheet-utils/plugins' },
    dbf: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    dif: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    slk: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    sylk: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    wks: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    wk1: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    wk3: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    wk4: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    '123': { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' },
    numbers: { plugin: 'legacyPlugin', importPath: 'spreadsheet-utils/plugins' }
};

/**
 * Register one or more format plugins globally.
 * @param {...Object} plugins
 */
export function registerPlugin(...plugins) {
    for (const plugin of plugins) {
        if (!plugin || typeof plugin !== 'object') continue;
        const name = plugin.name || `plugin_${Math.random().toString(36).slice(2)}`;
        pluginNames.add(name);
        pluginsMap.set(name, plugin);

        if (Array.isArray(plugin.extensions) && typeof plugin.parse === 'function') {
            for (const ext of plugin.extensions) {
                registeredParsers.set(ext.toLowerCase(), plugin.parse);
            }
        }
        if (typeof plugin.detect === 'function') {
            registeredSniffers.add(plugin.detect);
        }
        if (typeof plugin.write === 'function') {
            const writeExts = Array.isArray(plugin.writeExtensions)
                ? plugin.writeExtensions
                : (Array.isArray(plugin.extensions) ? plugin.extensions : []);
            for (const ext of writeExts) {
                registeredWriters.set(ext.toLowerCase(), plugin.write);
            }
        }
        if (plugin.writers && typeof plugin.writers === 'object') {
            for (const [ext, writerFn] of Object.entries(plugin.writers)) {
                if (typeof writerFn === 'function') {
                    registeredWriters.set(ext.toLowerCase(), writerFn);
                }
            }
        }
    }
}

/**
 * Unregister a plugin or reset registry.
 * @param {string|Object} [target] - Plugin name, object, or undefined to clear non-core plugins
 */
export function unregisterPlugin(target) {
    if (!target) {
        registeredParsers.clear();
        registeredWriters.clear();
        registeredSniffers.clear();
        pluginNames.clear();
        pluginsMap.clear();
        return;
    }
    const name = typeof target === 'string' ? target : target.name;
    const plugin = typeof target === 'string' ? pluginsMap.get(target) : target;

    if (name) {
        pluginNames.delete(name);
        pluginsMap.delete(name);
    }
    if (plugin && plugin.extensions) {
        for (const ext of plugin.extensions) {
            registeredParsers.delete(ext.toLowerCase());
            registeredWriters.delete(ext.toLowerCase());
        }
    }
    if (plugin && plugin.writeExtensions) {
        for (const ext of plugin.writeExtensions) {
            registeredWriters.delete(ext.toLowerCase());
        }
    }
    if (plugin && plugin.writers) {
        for (const ext of Object.keys(plugin.writers)) {
            registeredWriters.delete(ext.toLowerCase());
        }
    }
    if (plugin && plugin.detect) {
        registeredSniffers.delete(plugin.detect);
    }
}

/**
 * Get registered parser for a given extension.
 * @param {string} ext
 * @param {Map} [localRegistry]
 * @returns {Function|null}
 */
export function getParserForExtension(ext, localRegistry) {
    const key = (ext || '').toLowerCase().replace(/^\./, '');
    if (localRegistry && localRegistry.has(key)) {
        return localRegistry.get(key);
    }
    return registeredParsers.get(key) || null;
}

/**
 * Get registered writer for a given extension.
 * @param {string} ext
 * @param {Map} [localRegistry]
 * @returns {Function|null}
 */
export function getWriterForExtension(ext, localRegistry) {
    const key = (ext || '').toLowerCase().replace(/^\./, '');
    if (localRegistry && localRegistry.has(key)) {
        return localRegistry.get(key);
    }
    return registeredWriters.get(key) || null;
}

/**
 * Run registered content sniffers against raw bytes.
 * @param {Uint8Array} buffer
 * @param {Set} [localSniffers]
 * @returns {Promise<string|null>}
 */
export async function runSniffers(buffer, localSniffers) {
    const sniffers = localSniffers || registeredSniffers;
    for (const sniffer of sniffers) {
        try {
            const detected = await sniffer(buffer);
            if (detected) return detected;
        } catch {
            // Sniffer error, continue
        }
    }
    return null;
}

/**
 * Generate a friendly error explaining which plugin needs to be registered.
 * @param {string} ext
 * @param {string} [fileName]
 * @returns {Error}
 */
export function createMissingPluginError(ext, fileName) {
    const key = (ext || '').toLowerCase().replace(/^\./, '');
    const mapping = KNOWN_PLUGIN_MAPPINGS[key];
    if (mapping) {
        const err = new Error(
            `File format "${key}" requires the "${mapping.plugin}". ` +
            `Register it before reading or writing:\n\n` +
            `  import { ${mapping.plugin} } from '${mapping.importPath}';\n` +
            `  registerPlugin(${mapping.plugin});\n`
        );
        err.code = 'ERR_PLUGIN_REQUIRED';
        err.plugin = mapping.plugin;
        err.format = key;
        return err;
    }
    const err = new Error(`Unsupported file type: ${ext || fileName || '(unknown)'}`);
    err.code = 'ERR_UNSUPPORTED_FORMAT';
    return err;
}

/**
 * Check if a plugin is currently registered.
 * @param {string} name
 * @returns {boolean}
 */
export function isPluginRegistered(name) {
    return pluginNames.has(name);
}
