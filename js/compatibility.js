(function () {
    'use strict';

    if (!window.requestAnimationFrame) {
        window.requestAnimationFrame = function (callback) {
            return window.setTimeout(function () {
                callback(Date.now());
            }, 16);
        };
    }

    if (!window.cancelAnimationFrame) {
        window.cancelAnimationFrame = function (handle) {
            window.clearTimeout(handle);
        };
    }

    if (window.Element && !Element.prototype.matches) {
        Element.prototype.matches = Element.prototype.msMatchesSelector ||
            Element.prototype.webkitMatchesSelector ||
            function (selector) {
                var matches = (this.document || this.ownerDocument).querySelectorAll(selector);
                for (var index = 0; index < matches.length; index += 1) {
                    if (matches[index] === this) return true;
                }
                return false;
            };
    }

    if (window.Element && !Element.prototype.closest) {
        Element.prototype.closest = function (selector) {
            var element = this;
            while (element && element.nodeType === 1) {
                if (element.matches(selector)) return element;
                element = element.parentElement || element.parentNode;
            }
            return null;
        };
    }

    if (window.Element && !Element.prototype.replaceChildren) {
        Element.prototype.replaceChildren = function () {
            while (this.firstChild) this.removeChild(this.firstChild);
            for (var index = 0; index < arguments.length; index += 1) {
                this.appendChild(arguments[index]);
            }
        };
    }

    if (!Array.prototype.includes) {
        Array.prototype.includes = function (value, fromIndex) {
            var start = fromIndex || 0;
            if (start < 0) start = Math.max(this.length + start, 0);
            for (var index = start; index < this.length; index += 1) {
                if (this[index] === value || (value !== value && this[index] !== this[index])) return true;
            }
            return false;
        };
    }

    if (!String.prototype.includes) {
        String.prototype.includes = function (search, start) {
            return this.indexOf(search, start || 0) !== -1;
        };
    }

    if (!Object.entries) {
        Object.entries = function (object) {
            var entries = [];
            Object.keys(Object(object)).forEach(function (key) {
                entries.push([key, object[key]]);
            });
            return entries;
        };
    }

    if (!String.fromCodePoint) {
        String.fromCodePoint = function (codePoint) {
            if (codePoint <= 0xFFFF) return String.fromCharCode(codePoint);
            codePoint -= 0x10000;
            return String.fromCharCode(
                (codePoint >> 10) + 0xD800,
                (codePoint % 0x400) + 0xDC00
            );
        };
    }

    window.addPointerEventListener = function (target, type, listener, options) {
        var fallbackType = type;
        if (!window.PointerEvent) {
            fallbackType = {
                pointerdown: 'mousedown',
                pointermove: 'mousemove',
                pointerup: 'mouseup'
            }[type] || type;
        }
        target.addEventListener(fallbackType, listener, options);
        return fallbackType;
    };

    window.removePointerEventListener = function (target, type, listener, options) {
        var fallbackType = type;
        if (!window.PointerEvent) {
            fallbackType = {
                pointerdown: 'mousedown',
                pointermove: 'mousemove',
                pointerup: 'mouseup'
            }[type] || type;
        }
        target.removeEventListener(fallbackType, listener, options);
    };
}());
