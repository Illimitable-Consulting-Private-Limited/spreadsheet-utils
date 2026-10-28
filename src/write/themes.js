/**
 * Ready-to-Use Styling Themes for spreadsheet-utils.
 *
 * Provides cross-software consistent themes for XLSX, XLS, ODS, and HTML.
 * Renders identically across Microsoft Excel, LibreOffice Calc, and modern web browsers.
 */

export const THEMES = {
    ocean: {
        name: 'ocean',
        header: {
            backgroundColor: '#1B365D',
            color: '#FFFFFF',
            fontWeight: 'bold',
            borderColor: '#CBD5E1',
            borderStyle: 'thin'
        },
        alternateRow: {
            backgroundColor: '#F0F4F8'
        },
        borderColor: '#E2E8F0'
    },
    forest: {
        name: 'forest',
        header: {
            backgroundColor: '#1B4332',
            color: '#FFFFFF',
            fontWeight: 'bold',
            borderColor: '#D8F3DC',
            borderStyle: 'thin'
        },
        alternateRow: {
            backgroundColor: '#F2F9F4'
        },
        borderColor: '#E0EFE6'
    },
    slate: {
        name: 'slate',
        header: {
            backgroundColor: '#334155',
            color: '#FFFFFF',
            fontWeight: 'bold',
            borderColor: '#CBD5E1',
            borderStyle: 'thin'
        },
        alternateRow: {
            backgroundColor: '#F8FAFC'
        },
        borderColor: '#E2E8F0'
    },
    sunset: {
        name: 'sunset',
        header: {
            backgroundColor: '#9C4123',
            color: '#FFFFFF',
            fontWeight: 'bold',
            borderColor: '#FED7AA',
            borderStyle: 'thin'
        },
        alternateRow: {
            backgroundColor: '#FFF7ED'
        },
        borderColor: '#FEE2E2'
    },
    striped: {
        name: 'striped',
        header: {
            backgroundColor: '#475569',
            color: '#FFFFFF',
            fontWeight: 'bold',
            borderColor: '#CBD5E1',
            borderStyle: 'thin'
        },
        alternateRow: {
            backgroundColor: '#F1F5F9'
        },
        borderColor: '#E2E8F0'
    }
};

/**
 * Resolve theme name or custom theme object.
 *
 * @param {string|Object} theme
 * @returns {Object|null}
 */
export function getTheme(theme) {
    if (!theme) return null;
    if (typeof theme === 'string') {
        const key = theme.toLowerCase().trim();
        return THEMES[key] || null;
    }
    if (typeof theme === 'object') {
        return theme;
    }
    return null;
}

/**
 * Apply theme styling to 2D row grid.
 * Decorates header cells with header palette and applies alternating zebra background to data rows.
 *
 * @param {Array<Array<any>>} rows
 * @param {string|Object} themeInput
 * @param {Object} [options={}]
 * @returns {Array<Array<any>>}
 */
export function applyTheme(rows, themeInput, options = {}) {
    const theme = getTheme(themeInput);
    if (!theme || !Array.isArray(rows) || rows.length === 0) {
        return rows;
    }

    const hasHeader = options.hasHeader !== false;
    const headerStyle = theme.header || {};
    const alternateStyle = theme.alternateRow || {};
    const zebra = options.zebra !== false && Boolean(alternateStyle.backgroundColor);

    return rows.map((row, rIdx) => {
        if (!Array.isArray(row)) return row;

        const isHeader = hasHeader && rIdx === 0;
        const isAlternate = !isHeader && zebra && (hasHeader ? (rIdx % 2 === 0) : (rIdx % 2 === 1));

        return row.map((cell) => {
            if (isHeader) {
                if (cell && typeof cell === 'object' && !(cell instanceof Date)) {
                    return { ...headerStyle, ...cell };
                }
                return { value: cell !== undefined ? cell : null, ...headerStyle };
            }

            if (isAlternate) {
                if (cell && typeof cell === 'object' && !(cell instanceof Date)) {
                    return {
                        backgroundColor: alternateStyle.backgroundColor,
                        ...cell
                    };
                }
                return {
                    value: cell !== undefined ? cell : null,
                    backgroundColor: alternateStyle.backgroundColor
                };
            }

            return cell;
        });
    });
}
