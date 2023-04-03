const ADDON_ID = "declared-styles";

const ATTRIBUTES = {
    STYLES: {
        BACKGROUND_COLOR: {
            default: null,
        },
        BOLD: {
            default: false,
        },
        INDENT_END: {
            default: 0,
        },
        INDENT_FIRST_LINE: {
            default: 0,
        },
        INDENT_START: {
            default: 0,
        },
        ITALIC: {
            default: false,
        },
        STRIKETHROUGH: {
            default: false,
        },
        UNDERLINE: {
            default: false,
        },
    },
};

const HEADING_TYPES = [
    'HEADING1',
    'HEADING2',
    'HEADING3',
    'HEADING4',
    'HEADING5',
    'HEADING6',
    'TITLE',
    'SUBTITLE',
    'NORMAL',
];

const PAGE_SIZE_DETECT_EPSILON = 5.0;

// Page sizes in point units:
const PAGE_SIZES = {
    LETTER: [612.283, 790.866],
    TABLOID: [790.866, 1224.57],
    LEGAL: [612.283, 1009.13],
    STATEMENT: [396.85, 612.283],
    EXECUTIVE: [521.575, 756.85],
    FOLIO: [612.283, 935.433],
    A3: [841.89, 1190.55],
    A4: [595.276, 841.89],
    A5: [419.528, 595.276],
    B4: [708.661, 1000.63],
    B5: [498.898, 708.661],
};

const HandleStyleJsonApplyProxy = new Proxy({}, {
    get: (target, prop) => {
        return () => {
            const styleName = byteArrayToString(Utilities.base64DecodeWebSafe(base64EnsurePadding(prop)));
            const styleJson = storageStyleJsonGet(styleName);

            if (styleJson) {
                handleStyleJsonApply(styleJson)
            }
        };
    }
});

class UserPresentableError extends Error {
    constructor(message) {
        super(message);
        this.name = "UserPresentableError";
    }
}

function base64EnsurePadding(base64String) {
    const padding = '='.repeat((4 - base64String.length % 4) % 4);
    return base64String + padding;
}

function base64StripPadding(base64String) {
    const paddingIndex = base64String.indexOf('=');
    return paddingIndex === -1 ? base64String : base64String.slice(0, paddingIndex);
}

function byteArrayToString(byteArray) {
    var blob = Utilities.newBlob(byteArray, 'application/octet-stream');
    var string = blob.getDataAsString('UTF-8');
    return string;
}

function convertToPrettyJson(value) {
    return JSON.stringify(value, convertToPrettyJsonSortReplacer, 2);
}

function convertToPrettyJsonSortReplacer(key, value) {
    return (
        value instanceof Object && !(value instanceof Array) ?
            Object.keys(value)
                .sort()
                .reduce((sorted, key) => {
                    sorted[key] = value[key];
                    return sorted
                }, {}) :
            value
    );
}

function getStyle(options) {
    const body = DocumentApp.getActiveDocument().getBody();

    let style = {
        DOCUMENT: {
            ...getStylePageMargins(body, options),
            ...getStylePageSize(body, options),
        },
        META: {
            VERSION: "S1",
        },
        STYLES: (
            Object.fromEntries(
                HEADING_TYPES.map(
                    headingType => (
                        [headingType, transformHeadingAttributesToExternalStructure(body.getHeadingAttributes(DocumentApp.ParagraphHeading[headingType]))]
                    )
                )
            )
        ),
    };

    if (options?.stripDefaults) {
        style = stripStyleDefaults(style);
    }

    return style;
}

function isNumberEqual(a, b, epsilon = 0.00001) {
    return Math.abs(a - b) < epsilon;
}

function getStylePageMargins(body, options) {
    return {
        MARGIN_BOTTOM: body.getMarginBottom(),
        MARGIN_LEFT: body.getMarginLeft(),
        MARGIN_RIGHT: body.getMarginRight(),
        MARGIN_TOP: body.getMarginTop(),
    };
}

function getStylePageSize(body, options) {
    const pageWidth = body.getPageWidth();
    const pageHeight = body.getPageHeight();

    for (const [pageType, pageSize] of Object.entries(PAGE_SIZES)) {
        if (isNumberEqual(pageWidth, pageSize[0], PAGE_SIZE_DETECT_EPSILON) && isNumberEqual(pageHeight, pageSize[1], PAGE_SIZE_DETECT_EPSILON)) {
            return {
                PAGE_SIZE: pageType,
                PAGE_ORIENTATION: 'PORTRAIT',
            };
        } else if (isNumberEqual(pageWidth, pageSize[1], PAGE_SIZE_DETECT_EPSILON) && isNumberEqual(pageHeight, pageSize[0], PAGE_SIZE_DETECT_EPSILON)) {
            return {
                PAGE_SIZE: pageType,
                PAGE_ORIENTATION: 'LANDSCAPE',
            };
        }
    }

    return {
        PAGE_HEIGHT: pageHeight,
        PAGE_WIDTH: pageWidth,
    };
}

function handleOptionBooleanSet(name, value) {
    const options = storageObjectGet("options")
    options[name] = value;
    storageObjectSet("options", options);
}

function handleStyleJsonApply(styleJson, options) {
    handleStyleJsonSet(styleJson, options);

    const body = DocumentApp.getActiveDocument().getBody();

    const headerStyles = Object.fromEntries(
        HEADING_TYPES.map(
            headingType => (
                [
                    headingType,
                    body.getHeadingAttributes(DocumentApp.ParagraphHeading[headingType]),
                ]
            )
        )
    );

    const paragraphs = body.getParagraphs();

    for (const paragraph of paragraphs) {
        const paragraphHeaderStyleType = lookup(paragraph.getHeading(), DocumentApp.ParagraphHeading);
        if (paragraphHeaderStyleType != null) {
            paragraph.setAttributes(headerStyles[paragraphHeaderStyleType]);
        }
    }
}

function handleStyleJsonGet(options) {
    const style = getStyle(options);
    return convertToPrettyJson(style);
}

function handleStyleJsonSet(styleJson, options) {
    setStyle(JSON.parse(styleJson), options);
}

function lookup(value, targets) {
    for (const [targetKey, targetValue] of Object.entries(targets)) {
        if (Object.is(value, targetValue)) {
            return targetKey;
        }
    }

    return null;
}

function onOpen() {
    const ui = DocumentApp.getUi();
    const menu = ui.createMenu('JSON Styles');

    menu.addItem('Edit styles', 'showEditor')

    const styleIndex = storageStyleIndexGet();
    const styleNames = Object.keys(styleIndex).sort()

    if (styleNames.length !== 0) {
        const applyMenu = ui.createMenu("Apply style");

        for (const styleName of styleNames) {
            const encodedStyleName = base64StripPadding(Utilities.base64EncodeWebSafe(styleName));
            applyMenu.addItem(styleName, `HandleStyleJsonApplyProxy.${encodedStyleName}`);//handlerName);
        }

        menu.addSubMenu(applyMenu);
    }

    menu.addToUi();
}

function setStyle(style, options) {
    const body = DocumentApp.getActiveDocument().getBody();

    if (style?.STYLES != null) {
        Object.entries(style.STYLES).forEach(
            ([headingType, headingValue]) => {
                body.setHeadingAttributes(
                    DocumentApp.ParagraphHeading[headingType],
                    transformHeadingAttributesToInternalStructure(headingValue),
                );
            }
        )
    }

    if (style?.DOCUMENT != null) {
        setStylePageMargins(style.DOCUMENT, body, options);
        setStylePageSize(style.DOCUMENT, body, options);
    }
}

function setStylePageMargins(documentStyles, body, options) {
    if (documentStyles?.MARGIN_BOTTOM) {
        body.setMarginBottom(documentStyles.MARGIN_BOTTOM);
    }
    if (documentStyles?.MARGIN_LEFT) {
        body.setMarginLeft(documentStyles.MARGIN_LEFT);
    }
    if (documentStyles?.MARGIN_RIGHT) {
        body.setMarginRight(documentStyles.MARGIN_RIGHT);
    }
    if (documentStyles?.MARGIN_TOP) {
        body.setMarginTop(documentStyles.MARGIN_TOP);
    }
};

function setStylePageSize(documentStyles, body, options) {
    if (documentStyles?.PAGE_SIZE && !(PAGE_SIZES.hasOwnProperty(documentStyles.PAGE_SIZE))) {
        throw new UserPresentableError(`unrecognized page size: ${documentStyles.PAGE_SIZE}`)
    }

    const size = PAGE_SIZES?.[documentStyles?.PAGE_SIZE];
    const PAGE_ORIENTATION = documentStyles?.PAGE_ORIENTATION;

    const pageWidth = (
        PAGE_ORIENTATION === 'LANDSCAPE' ? (size?.[1] ?? documentStyles?.PAGE_WIDTH) : (size?.[0] ?? documentStyles?.PAGE_WIDTH)
    );

    const pageHeight = (
        PAGE_ORIENTATION === 'LANDSCAPE' ? (size?.[0] ?? documentStyles?.PAGE_HEIGHT) : (size?.[1] ?? documentStyles?.PAGE_HEIGHT)
    );

    if (pageWidth) {
        body.setPageWidth(pageWidth);
    }

    if (pageHeight) {
        body.setPageHeight(pageHeight);
    }
}

function showEditor() {
    var html = (
        HtmlService
            .createHtmlOutputFromFile('editor')
            .setTitle('Declared Styles')
    );

    DocumentApp.getUi().showSidebar(html);
}

function storageObjectGet(key) {
    const properties = PropertiesService.getUserProperties();
    return JSON.parse(properties.getProperty(key) || "{}");
}

function storageObjectSet(key, value) {
    const properties = PropertiesService.getUserProperties();
    properties.setProperty(key, JSON.stringify(value));
}

function storageStyleJsonDelete(styleName) {
    const styleIndex = storageStyleIndexGet();
    delete styleIndex[styleName];
    storageStyleIndexSet(styleIndex);

    const styleKey = `style:${styleName}`;
    const properties = PropertiesService.getUserProperties();
    properties.deleteProperty(styleKey);

    return { styleIndex };
}

function storageStyleJsonGet(styleName) {
    const key = `style:${styleName}`;

    const properties = PropertiesService.getUserProperties();
    const styleJsonCompressed = properties.getProperty(key);

    if (styleJsonCompressed) {
        const styleJson = LZString.decompressFromUTF16(styleJsonCompressed);
        return styleJson;
    } else {
        return null;
    }
}

function storageStyleIndexGet() {
    const properties = PropertiesService.getUserProperties();
    const styleIndex = JSON.parse(properties.getProperty("styleIndex") || "{}");
    return styleIndex;
}

function storageStyleIndexSet(styleIndex) {
    storageObjectSet("styleIndex", styleIndex);
}

function storageStyleJsonSet(styleName, styleJson) {
    var properties = PropertiesService.getUserProperties();

    const styleIndex = storageStyleIndexGet();
    styleIndex[styleName] = true;

    const styleKey = `style:${styleName}`;
    const styleJsonCompressed = LZString.compressToUTF16(styleJson);

    properties.setProperties({
        styleIndex: JSON.stringify(styleIndex),
        [styleKey]: styleJsonCompressed,
    });

    return { styleIndex };
}

function stripStyleDefaults(style) {
    return Object.fromEntries(
        Object.entries(style).map(
            ([attributeCategoryName, attributeGroups]) => (
                [
                    attributeCategoryName,
                    (
                        (ATTRIBUTES?.[attributeCategoryName] != null)
                            ? Object.fromEntries(
                                Object.entries(attributeGroups).map(
                                    ([attributeGroupName, attributeGroup]) => (
                                        [
                                            attributeGroupName,
                                            Object.fromEntries(
                                                Object.entries(attributeGroup).filter(
                                                    ([attributeName, attributeValue]) => (attributeValue !== ATTRIBUTES?.[attributeCategoryName]?.[attributeName]?.default)
                                                )
                                            )
                                        ]
                                    )
                                ).filter(
                                    ([attributeGroupName, attributeGroup]) => (attributeGroup && Object.keys(attributeGroup).length > 0)
                                )
                            ) : attributeGroups
                    )
                ]
            )
        )
    );
}

function transformHeadingAttributesToExternalStructure(headingAttributes) {
    const result = Object.assign({}, headingAttributes);
    if (headingAttributes.HORIZONTAL_ALIGNMENT != null) {
        result.HORIZONTAL_ALIGNMENT = lookup(headingAttributes.HORIZONTAL_ALIGNMENT, DocumentApp.HorizontalAlignment);
    }
    if (headingAttributes.VERTICAL_ALIGNMENT != null) {
        result.VERTICAL_ALIGNMENT = lookup(headingAttributes.VERTICAL_ALIGNMENT, DocumentApp.VerticalAlignment);
    }
    return result;
}

function transformHeadingAttributesToInternalStructure(headingObject) {
    const result = Object.assign({}, headingObject);
    if (headingObject.HORIZONTAL_ALIGNMENT != null) {
        result.HORIZONTAL_ALIGNMENT = DocumentApp.HorizontalAlignment[headingObject.HORIZONTAL_ALIGNMENT];
    }
    if (headingObject.VERTICAL_ALIGNMENT != null) {
        result.VERTICAL_ALIGNMENT = DocumentApp.VerticalAlignment[headingObject.VERTICAL_ALIGNMENT];
    }
    return result;
}
