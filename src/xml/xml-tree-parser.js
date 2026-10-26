/**
 * Parse HTML/XML string into a JSON tree structure
 * @param {string} html - HTML or XML string to parse
 * @param {Object} options - Parser options
 * @param {string[]} options.ignore - Array of tag names to ignore (e.g., ['script', 'style'])
 * @returns {Object} Parsed JSON tree
 */
export default function parser(html, options) {
    // Handle null/undefined options
    if (!options || typeof options !== 'object') {
        options = {};
    }

    // Normalize ignore list to lowercase for case-insensitive matching
    const ignoreTags = new Set((options.ignore || []).map(tag => tag.toLowerCase()));
    /**
     * Check if is a self-closing tag
     * @param {string} type - Tag name
     * @returns {boolean}
     */
    function isSelfClosing(type) {
        if (!type) {
            return false;
        }
        // List of self-closing or void HTML elements
        const selfClosingTags = [
            'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'source', 'track', 'wbr'
        ];
        return selfClosingTags.includes(type.toLowerCase());
    }

    /**
     * Create a text node and add it to current node's children
     * @param {Object} tag - Text node properties
     */
    const createTextNode = function(tag) {
        if (!this.current.children) {
            this.current.children = [];
        }

        this.current.children.push({
            type: '#text',
            props: [tag],
        });
    }

    /**
     * Find the parent node by tag name (traverses up using control.stack)
     * @param {string} type - Tag name to find
     * @returns {Object|undefined}
     */
    const findParentByTagName = function(type) {
        // Search backwards through the stack to find matching opening tag
        for (let i = this.stack.length - 1; i >= 0; i--) {
            if (this.stack[i].type === type) {
                return this.stack[i];
            }
        }
        return undefined;
    }

    /**
     * Handle the text node creation
     */
    const commitText = function() {
        if (typeof(this.text) !== 'undefined') {
            // Skip if we're inside an ignored tag
            if (this.ignoreDepth === 0) {
                // Preserve whitespace as-is
                let text = this.text;

                if (text) {
                    createTextNode.call(this, { name: 'textContent', value: text });
                }
            }
            delete this.text;
        }
    }

    /**
     * Handle comment node creation
     */
    const commitComments = function() {
        if (typeof(this.comments) !== 'undefined') {
            // Skip if we're inside an ignored tag
            if (this.ignoreDepth === 0) {
                let comments = this.comments;
                if (comments) {
                    comments = comments
                        .replace('<!--', '')
                        .replace('-->', '')

                    if (!this.current.children) {
                        this.current.children = [];
                    }

                    this.current.children.push({
                        type: '#comments',
                        props: [{ name: 'text', value: comments }],
                    });
                }
            }
            delete this.comments;
        }
    }

    /**
     * Save the attribute to the tag
     */
    const commitAttribute = function() {
        if (this.tag.attributeName) {
            // Commit any current attribute
            if (!this.tag.props) {
                this.tag.props = [];
            }

            let k = this.tag.attributeName;
            let v = this.tag.attributeValue;

            if (typeof(v) === 'undefined') {
                v = k;
            }

            this.tag.props.push({
                name: k,
                value: v,
            });

            // Clean up temporary properties
            delete this.tag.attributeName;
            delete this.tag.attributeValue;

            if (this.tag.attributeIsReadyToClose) {
                delete this.tag.attributeIsReadyToClose;
            }
        }
    }

    /**
     * Actions controller
     * @param {Object} control - Parser control object
     * @param {string} char - Current character
     */
    const actions = function(control, char) {
        const method = control.action || 'text';
        if (typeof actions[method] === 'function') {
            actions[method].call(control, char);
        }
    }

    /**
     * Process a new tag
     * @param char
     */
    actions.processTag = function(char) {
        // Just to check if there are any text to commit
        commitText.call(this);

        // Process the tag
        if (char === '<') {
            // Create new tag
            this.tag = {
                type: ''
            };
        } else if (char.match(/[a-zA-Z0-9-:]/)) {
            // Tag name (including colons for XML namespaces)
            this.tag.type += char;
        } else {
            // Finished with tag name, move to attribute handling
            this.action = 'attributeName';
        }
    }

    /**
     * Handle tag closing
     * @param char
     */
    actions.closeTag = function(char) {
        // Make sure to commit attribute
        commitAttribute.call(this);
        // Close the tag
        if (char === '>') {
            // Check if this tag should be ignored
            const shouldIgnore = ignoreTags.has(this.tag.type.toLowerCase());

            // Handle closing tags for ignored elements
            if (this.tag.closingTag && shouldIgnore && this.ignoreDepth > 0) {
                this.ignoreDepth--;
                this.tag = null;
                this.action = 'text';
                return;
            }

            // If we're inside an ignored tag, skip processing
            if (this.ignoreDepth > 0) {
                this.tag = null;
                this.action = 'text';
                return;
            }

            // Opening tag for ignored element - increment depth and skip
            if (shouldIgnore && !this.tag.closingTag) {
                this.ignoreDepth++;
                this.tag = null;
                this.action = 'text';
                return;
            }

            // Get the new parent
            if (isSelfClosing(this.tag.type)) {
                // Push new tag to the current
                if (!this.current.children) {
                    this.current.children = [];
                }
                this.current.children.push(this.tag);
            } else if (this.tag.closingTag) {
                // Need to find the parent on the chain
                const parentNode = findParentByTagName.call(this, this.tag.type);
                if (parentNode) {
                    // Pop stack until we find the matching tag
                    while (this.stack.length > 0 && this.stack[this.stack.length - 1] !== parentNode) {
                        this.stack.pop();
                    }
                    // Pop the matched tag itself
                    if (this.stack.length > 0) {
                        this.stack.pop();
                    }
                    // Current is now the top of stack (or root if empty)
                    this.current = this.stack.length > 0 ? this.stack[this.stack.length - 1] : this.root;
                }
            } else {
                // Store the parent before updating current
                if (this.tag.closing) {
                    // Self-closing tag like <br />
                    if (!this.current.children) {
                        this.current.children = [];
                    }
                    this.current.children.push(this.tag);
                } else {
                    // Opening tag - push to current's children
                    if (!this.current.children) {
                        this.current.children = [];
                    }
                    this.current.children.push(this.tag);

                    // Push to stack and update current
                    this.stack.push(this.tag);
                    this.current = this.tag;
                }
            }

            // Clean up temporary properties
            delete this.tag.insideQuote;
            delete this.tag.closingTag;
            delete this.tag.closing;
            delete this.tag.locked;
            // Finalize tag
            this.tag = null;
            // New action
            this.action = 'text';
        } else if (!this.tag.locked) {
            if (char === '/') {
                if (!this.tag.type) {
                    // This is a closing tag
                    this.tag.closingTag = true;
                }
                // Closing character is found
                this.tag.closing = true;
            } else if (char.match(/[a-zA-Z0-9-:]/)) {
                // If is a closing tag, get the tag name (including colons for XML namespaces)
                if (this.tag.closingTag) {
                    this.tag.type += char;
                }
            } else {
                // Wait to the closing sign
                if (this.tag.type) {
                    this.tag.locked = true;
                }
            }
        }
    }

    actions.attributeName = function(char) {
        // There is another attribute to commit
        if (this.tag.attributeIsReadyToClose) {
            commitAttribute.call(this);
        }

        // Build attribute name
        if (char.match(/[a-zA-Z0-9-:]/)) {
            if (!this.tag.attributeName) {
                this.tag.attributeName = '';
            }
            this.tag.attributeName += char;
        } else if (char === '=') {
            // Move to attribute value
            if (this.tag.attributeName) {
                this.action = 'attributeValue';
                delete this.tag.attributeIsReadyToClose;
            }
        } else if (char.match(/\s/)) {
            if (this.tag.attributeName) {
                this.tag.attributeIsReadyToClose = true;
            }
        }
    };

    actions.attributeValue = function(char) {
        if (!this.tag.attributeValue) {
            this.tag.attributeValue = '';
        }

        if (char === '"' || char === "'") {
            if (this.tag.insideQuote) {
                if (this.tag.insideQuote === char) {
                    this.tag.insideQuote = false;
                } else {
                    this.tag.attributeValue += char;
                }
            } else {
                this.tag.insideQuote = char;
            }
        } else {
            // Inside quotes, keep appending to the attribute value
            if (this.tag.insideQuote) {
                if (this.tag.attributeValue) {
                    this.tag.attributeValue += char;
                } else {
                    this.tag.attributeValue = char;
                }
            } else if (typeof(char) === 'string' && char.match(/\s/)) {
                if (this.tag.attributeValue) {
                    this.action = 'attributeName';
                }
                this.tag.attributeIsReadyToClose = true;
            } else {
                if (this.tag.attributeIsReadyToClose) {
                    this.action = 'attributeName';
                    actions.attributeName.call(this, char);
                } else {
                    if (this.tag.attributeValue) {
                        this.tag.attributeValue += char;
                    } else {
                        this.tag.attributeValue = char;
                    }
                }
            }
        }
    }

    actions.text = function(char) {
        // Normal text processing
        if (!this.text) {
            this.text = '';
        }
        this.text += char; // Keep appending to text content
    }

    actions.comments = function(char) {
        if (!this.comments) {
            this.comments = '';
        }
        this.comments += char;

        if (this.comments.endsWith('-->')) {
            commitComments.call(this);
            this.action = 'text';
        }
    }

    const result = { type: 'template' };
    const control = {
        root: result,
        current: result,
        stack: [],  // Stack to track open tags
        action: 'text',
        ignoreDepth: 0,  // Track depth of ignored tags
    };

    // Input validation
    if (typeof html !== 'string') {
        throw new TypeError('HTML input must be a string');
    }

    // Main loop to process the HTML string
    for (let i = 0; i < html.length; i++) {
        // Current char
        let char = html[i];

        if (control.action === 'text' && char === '<' && html[i+1] === '!' && html[i+2] === '-' && html[i+3] === '-') {
            control.action = 'comments';
        }

        if (control.action !== 'comments') {
            // Global control logic
            if (control.tag) {
                if (char === '>' || char === '/') {
                    // End of tag, commit any attributes and go back to text parsing
                    if (!control.tag.insideQuote) {
                        control.action = 'closeTag';
                    }
                }
            } else {
                if (char === '<') {
                    control.action = 'processTag';
                }
            }
        }

        // Execute action
        actions(control, char);
    }

    // Handle any remaining text
    commitText.call(control);

    // Determine the final result
    let finalResult;
    if (result.children) {
        // If there's only one root element, return it directly
        if (result.children.length === 1) {
            finalResult = result.children[0];
        }
        // If there are multiple root elements, return the template wrapper
        else if (result.children.length > 1) {
            finalResult = result;
        }
    }

    return finalResult;
}

export { parser };
