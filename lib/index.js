'use strict';

Object.defineProperty(exports, '__esModule', { value: true });

var unistUtilVisit = require('unist-util-visit');
var docx = require('docx');
var unifiedLatexUtilParse = require('@unified-latex/unified-latex-util-parse');

/**
 * @internal
 */
const unreachable = (_) => {
    throw new Error("unreachable");
};
/**
 * @internal
 */
function invariant(cond, message) {
    if (!cond)
        throw new Error(message);
}

const hasSquareBrackets = (arg) => {
    return !!arg && arg.openMark === "[" && arg.closeMark === "]";
};
const hasCurlyBrackets = (arg) => {
    return !!arg && arg.openMark === "{" && arg.closeMark === "}";
};
const mapString = (s) => new docx.MathRun(s);
const mapMacro = (n, runs) => {
    var _a, _b, _c, _d, _e, _f, _g, _h;
    switch (n.content) {
        case "#":
            return mapString("#");
        case "$":
            return mapString("$");
        case "%":
            return mapString("%");
        case "&":
            return mapString("&");
        case "textasciitilde":
            return mapString("~");
        case "textasciicircum":
            return mapString("^");
        case "textbackslash":
            return mapString("∖");
        case "{":
            return mapString("{");
        case "}":
            return mapString("}");
        case "textbar":
            return mapString("|");
        case "textless":
            return mapString("<");
        case "textgreater":
            return mapString(">");
        case "neq":
            return mapString("≠");
        case "sim":
            return mapString("∼");
        case "simeq":
            return mapString("≃");
        case "approx":
            return mapString("≈");
        case "fallingdotseq":
            return mapString("≒");
        case "risingdotseq":
            return mapString("≓");
        case "equiv":
            return mapString("≡");
        case "geq":
            return mapString("≥");
        case "geqq":
            return mapString("≧");
        case "leq":
            return mapString("≤");
        case "leqq":
            return mapString("≦");
        case "gg":
            return mapString("≫");
        case "ll":
            return mapString("≪");
        case "times":
            return mapString("×");
        case "div":
            return mapString("÷");
        case "pm":
            return mapString("±");
        case "mp":
            return mapString("∓");
        case "oplus":
            return mapString("⊕");
        case "ominus":
            return mapString("⊖");
        case "otimes":
            return mapString("⊗");
        case "oslash":
            return mapString("⊘");
        case "circ":
            return mapString("∘");
        case "cdot":
            return mapString("⋅");
        case "bullet":
            return mapString("∙");
        case "ltimes":
            return mapString("⋉");
        case "rtimes":
            return mapString("⋊");
        case "in":
            return mapString("∈");
        case "ni":
            return mapString("∋");
        case "notin":
            return mapString("∉");
        case "subset":
            return mapString("⊂");
        case "supset":
            return mapString("⊃");
        case "subseteq":
            return mapString("⊆");
        case "supseteq":
            return mapString("⊇");
        case "nsubseteq":
            return mapString("⊈");
        case "nsupseteq":
            return mapString("⊉");
        case "subsetneq":
            return mapString("⊊");
        case "supsetneq":
            return mapString("⊋");
        case "cap":
            return mapString("∩");
        case "cup":
            return mapString("∪");
        case "emptyset":
            return mapString("∅");
        case "infty":
            return mapString("∞");
        case "partial":
            return mapString("∂");
        case "aleph":
            return mapString("ℵ");
        case "hbar":
            return mapString("ℏ");
        case "wp":
            return mapString("℘");
        case "Re":
            return mapString("ℜ");
        case "Im":
            return mapString("ℑ");
        case "alpha":
            return mapString("α");
        case "beta":
            return mapString("β");
        case "gamma":
            return mapString("γ");
        case "delta":
            return mapString("δ");
        case "epsilon":
            return mapString("ϵ");
        case "zeta":
            return mapString("ζ");
        case "eta":
            return mapString("η");
        case "theta":
            return mapString("θ");
        case "iota":
            return mapString("ι");
        case "kappa":
            return mapString("κ");
        case "lambda":
            return mapString("λ");
        case "eta":
            return mapString("η");
        case "mu":
            return mapString("μ");
        case "nu":
            return mapString("ν");
        case "xi":
            return mapString("ξ");
        case "pi":
            return mapString("π");
        case "rho":
            return mapString("ρ");
        case "sigma":
            return mapString("σ");
        case "tau":
            return mapString("τ");
        case "upsilon":
            return mapString("υ");
        case "phi":
            return mapString("ϕ");
        case "chi":
            return mapString("χ");
        case "psi":
            return mapString("ψ");
        case "omega":
            return mapString("ω");
        case "varepsilon":
            return mapString("ε");
        case "vartheta":
            return mapString("ϑ");
        case "varrho":
            return mapString("ϱ");
        case "varsigma":
            return mapString("ς");
        case "varphi":
            return mapString("φ");
        case "Gamma":
            return mapString("Γ");
        case "Delta":
            return mapString("Δ");
        case "Theta":
            return mapString("Θ");
        case "Lambda":
            return mapString("Λ");
        case "Xi":
            return mapString("Ξ");
        case "Pi":
            return mapString("Π");
        case "Sigma":
            return mapString("Σ");
        case "Upsilon":
            return mapString("Υ");
        case "Phi":
            return mapString("Φ");
        case "Psi":
            return mapString("Ψ");
        case "Omega":
            return mapString("Ω");
        case "newline":
        case "\\":
            // line break
            return false;
        case "^": {
            const prev = runs.pop();
            if (!prev)
                break;
            return new docx.MathSuperScript({
                children: [prev],
                superScript: mapGroup((_c = (_b = (_a = n.args) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b.content) !== null && _c !== void 0 ? _c : []),
            });
        }
        case "_": {
            const prev = runs.pop();
            if (!prev)
                break;
            return new docx.MathSubScript({
                children: [prev],
                subScript: mapGroup((_f = (_e = (_d = n.args) === null || _d === void 0 ? void 0 : _d[0]) === null || _e === void 0 ? void 0 : _e.content) !== null && _f !== void 0 ? _f : []),
            });
        }
        case "hat":
            // TODO: implement
            break;
        case "widehat":
            // TODO: implement
            break;
        case "sum": {
            // TODO: support superscript and subscript
            return new docx.MathSum({
                children: [],
            });
        }
        case "int":
            return mapString("∫");
        case "frac":
        case "tfrac":
        case "dfrac": {
            const args = (_g = n.args) !== null && _g !== void 0 ? _g : [];
            if (args.length === 2 &&
                hasCurlyBrackets(args[0]) &&
                hasCurlyBrackets(args[1])) {
                return new docx.MathFraction({
                    numerator: mapGroup(args[0].content),
                    denominator: mapGroup(args[1].content),
                });
            }
            break;
        }
        case "sqrt": {
            const args = (_h = n.args) !== null && _h !== void 0 ? _h : [];
            if (args.length === 1 && hasCurlyBrackets(args[0])) {
                return new docx.MathRadical({
                    children: mapGroup(args[0].content),
                });
            }
            if (args.length === 2 &&
                hasSquareBrackets(args[0]) &&
                hasCurlyBrackets(args[1])) {
                return new docx.MathRadical({
                    children: mapGroup(args[1].content),
                    degree: mapGroup(args[0].content),
                });
            }
            break;
        }
    }
    return mapString(n.content);
};
const mapGroup = (nodes) => {
    const group = [];
    for (const c of nodes) {
        group.push(...(mapNode(c, group) || []));
    }
    return group;
};
const mapNode = (n, runs) => {
    switch (n.type) {
        case "root":
            break;
        case "string":
            return [mapString(n.content)];
        case "whitespace":
            break;
        case "parbreak":
            break;
        case "comment":
            break;
        case "macro":
            const run = mapMacro(n, runs);
            if (!run) {
                // line break
                return false;
            }
            else {
                return [run];
            }
        case "environment":
        case "mathenv":
            break;
        case "verbatim":
            break;
        case "inlinemath":
            break;
        case "displaymath":
            break;
        case "group":
            return mapGroup(n.content);
        case "verb":
            break;
        default:
            unreachable();
    }
    return [];
};
/**
 * @internal
 */
const parseLatex = (value) => {
    const parsed = unifiedLatexUtilParse.parseMath(value);
    const paragraphs = [[]];
    let runs = paragraphs[0];
    for (const n of parsed) {
        const res = mapNode(n, runs);
        if (!res) {
            // line break
            paragraphs.push((runs = []));
        }
        else {
            runs.push(...res);
        }
    }
    return paragraphs;
};

const ORDERED_LIST_REF = "ordered";
const INDENT = 0.5;
const DEFAULT_NUMBERINGS = [
    {
        level: 0,
        format: docx.LevelFormat.DECIMAL,
        text: "%1.",
        alignment: docx.AlignmentType.START,
    },
    {
        level: 1,
        format: docx.LevelFormat.DECIMAL,
        text: "%2.",
        alignment: docx.AlignmentType.START,
        style: {
            paragraph: {
                indent: { start: docx.convertInchesToTwip(INDENT * 1) },
            },
        },
    },
    {
        level: 2,
        format: docx.LevelFormat.DECIMAL,
        text: "%3.",
        alignment: docx.AlignmentType.START,
        style: {
            paragraph: {
                indent: { start: docx.convertInchesToTwip(INDENT * 2) },
            },
        },
    },
    {
        level: 3,
        format: docx.LevelFormat.DECIMAL,
        text: "%4.",
        alignment: docx.AlignmentType.START,
        style: {
            paragraph: {
                indent: { start: docx.convertInchesToTwip(INDENT * 3) },
            },
        },
    },
    {
        level: 4,
        format: docx.LevelFormat.DECIMAL,
        text: "%5.",
        alignment: docx.AlignmentType.START,
        style: {
            paragraph: {
                indent: { start: docx.convertInchesToTwip(INDENT * 4) },
            },
        },
    },
    {
        level: 5,
        format: docx.LevelFormat.DECIMAL,
        text: "%6.",
        alignment: docx.AlignmentType.START,
        style: {
            paragraph: {
                indent: { start: docx.convertInchesToTwip(INDENT * 5) },
            },
        },
    },
];
const mdastToDocx = async (node, { output = "buffer", title, subject, creator, keywords, description, lastModifiedBy, revision, styles, background, tableStyle, }, images) => {
    const { nodes, footnotes } = convertNodes(node.children, {
        deco: {},
        images,
        indent: 0,
        tableStyle
    });
    const doc = new docx.Document({
        title,
        subject,
        creator,
        keywords,
        description,
        lastModifiedBy,
        revision,
        styles,
        background,
        footnotes,
        sections: [{ children: nodes }],
        numbering: {
            config: [
                {
                    reference: ORDERED_LIST_REF,
                    levels: DEFAULT_NUMBERINGS,
                },
            ],
        },
    });
    switch (output) {
        case "buffer":
            const bufOut = await docx.Packer.toBuffer(doc);
            // feature detection instead of environment detection, but if Buffer exists
            // it's probably Node. If not, return the Uint8Array that JSZip returns
            // when it doesn't detect a Node environment.
            return typeof Buffer === "function" ? Buffer.from(bufOut) : bufOut;
        case "blob":
            return docx.Packer.toBlob(doc);
    }
};
const convertNodes = (nodes, ctx) => {
    const results = [];
    let footnotes = {};
    for (const node of nodes) {
        switch (node.type) {
            case "paragraph":
                results.push(buildParagraph(node, ctx));
                break;
            case "heading":
                results.push(buildHeading(node, ctx));
                break;
            case "thematicBreak":
                results.push(buildThematicBreak());
                break;
            case "blockquote":
                results.push(...buildBlockquote(node, ctx));
                break;
            case "list":
                results.push(...buildList(node, ctx));
                break;
            case "listItem":
                invariant(false, "unreachable");
            case "table":
                results.push(buildTable(node, ctx));
                break;
            case "tableRow":
                invariant(false, "unreachable");
            case "tableCell":
                invariant(false, "unreachable");
            case "html":
                results.push(buildHtml(node));
                break;
            case "code":
                results.push(buildCode(node));
                break;
            case "yaml":
                // FIXME: unimplemented
                break;
            case "toml":
                // FIXME: unimplemented
                break;
            case "definition":
                // FIXME: unimplemented
                break;
            case "footnoteDefinition":
                footnotes[node.identifier] = buildFootnoteDefinition(node, ctx);
                break;
            case "text":
                results.push(buildText(node.value, ctx.deco));
                break;
            case "emphasis":
            case "strong":
            case "delete": {
                const { type, children } = node;
                const { nodes } = convertNodes(children, {
                    ...ctx,
                    deco: { ...ctx.deco, [type]: true },
                });
                results.push(...nodes);
                break;
            }
            case "inlineCode":
                // FIXME: transform to text for now
                results.push(buildText(node.value, ctx.deco));
                break;
            case "break":
                results.push(buildBreak());
                break;
            case "link":
                results.push(buildLink(node, ctx));
                break;
            case "image":
                results.push(buildImage(node, ctx.images));
                break;
            case "linkReference":
                // FIXME: unimplemented
                break;
            case "imageReference":
                // FIXME: unimplemented
                break;
            case "footnote":
                results.push(buildFootnote(node, ctx));
                break;
            case "footnoteReference":
                // do we need context here?
                results.push(buildFootnoteReference(node));
                break;
            case "math":
                results.push(...buildMath(node));
                break;
            case "inlineMath":
                results.push(buildInlineMath(node));
                break;
            default:
                unreachable();
                break;
        }
    }
    return {
        nodes: results,
        footnotes,
    };
};
const buildParagraph = ({ children }, ctx) => {
    const list = ctx.list;
    const { nodes } = convertNodes(children, ctx);
    if (list && list.checked != null) {
        nodes.unshift(new docx.CheckBox({
            checked: list.checked,
            checkedState: { value: "2611" },
            uncheckedState: { value: "2610" },
        }));
    }
    return new docx.Paragraph({
        children: nodes,
        indent: ctx.indent > 0
            ? {
                start: docx.convertInchesToTwip(INDENT * ctx.indent),
            }
            : undefined,
        ...(list &&
            (list.ordered
                ? {
                    numbering: {
                        reference: ORDERED_LIST_REF,
                        level: list.level,
                    },
                }
                : {
                    bullet: {
                        level: list.level,
                    },
                })),
    });
};
const buildHeading = ({ children, depth }, ctx) => {
    let heading;
    switch (depth) {
        case 1:
            heading = docx.HeadingLevel.TITLE;
            break;
        case 2:
            heading = docx.HeadingLevel.HEADING_1;
            break;
        case 3:
            heading = docx.HeadingLevel.HEADING_2;
            break;
        case 4:
            heading = docx.HeadingLevel.HEADING_3;
            break;
        case 5:
            heading = docx.HeadingLevel.HEADING_4;
            break;
        case 6:
            heading = docx.HeadingLevel.HEADING_5;
            break;
    }
    const { nodes } = convertNodes(children, ctx);
    return new docx.Paragraph({
        heading,
        children: nodes,
    });
};
const buildThematicBreak = (_) => {
    return new docx.Paragraph({
        thematicBreak: true,
    });
};
const buildBlockquote = ({ children }, ctx) => {
    const { nodes } = convertNodes(children, { ...ctx, indent: ctx.indent + 1 });
    return nodes;
};
const buildList = ({ children, ordered, start: _start, spread: _spread }, ctx) => {
    const list = {
        level: ctx.list ? ctx.list.level + 1 : 0,
        ordered: !!ordered,
    };
    return children.flatMap((item) => {
        return buildListItem(item, {
            ...ctx,
            list,
        });
    });
};
const buildListItem = ({ children, checked, spread: _spread }, ctx) => {
    const { nodes } = convertNodes(children, {
        ...ctx,
        ...(ctx.list && { list: { ...ctx.list, checked: checked !== null && checked !== void 0 ? checked : undefined } }),
    });
    return nodes;
};
const buildTable = ({ children, align }, ctx) => {
    var _a;
    const cellAligns = align === null || align === void 0 ? void 0 : align.map((a) => {
        switch (a) {
            case "left":
                return docx.AlignmentType.LEFT;
            case "right":
                return docx.AlignmentType.RIGHT;
            case "center":
                return docx.AlignmentType.CENTER;
            default:
                return docx.AlignmentType.LEFT;
        }
    });
    return new docx.Table({
        ...(_a = ctx.tableStyle) === null || _a === void 0 ? void 0 : _a.options,
        rows: children.map((r, i) => {
            var _a, _b;
            const style = i === 0 ? (_a = ctx.tableStyle) === null || _a === void 0 ? void 0 : _a.header : (_b = ctx.tableStyle) === null || _b === void 0 ? void 0 : _b.body;
            return buildTableRow(r, ctx, cellAligns, style);
        }),
    });
};
const buildTableRow = ({ children }, ctx, cellAligns, style) => {
    return new docx.TableRow({
        ...style === null || style === void 0 ? void 0 : style.row,
        children: children.map((c, i) => {
            return buildTableCell(c, ctx, cellAligns === null || cellAligns === void 0 ? void 0 : cellAligns[i], style);
        }),
    });
};
const buildTableCell = ({ children }, ctx, align, style) => {
    const { nodes } = convertNodes(children, ctx);
    return new docx.TableCell({
        ...style === null || style === void 0 ? void 0 : style.cell,
        children: [
            new docx.Paragraph({
                ...style === null || style === void 0 ? void 0 : style.paragraph,
                alignment: align,
                children: nodes,
            }),
        ],
    });
};
const buildHtml = ({ value }) => {
    // FIXME: transform to text for now
    return new docx.Paragraph({
        children: [buildText(value, {})],
    });
};
const buildCode = ({ value, lang: _lang, meta: _meta }) => {
    // FIXME: transform to text for now
    return new docx.Paragraph({
        children: [buildText(value, {})],
    });
};
const buildMath = ({ value }) => {
    return parseLatex(value).map((runs) => new docx.Paragraph({
        children: [
            new docx.Math({
                children: runs,
            }),
        ],
    }));
};
const buildInlineMath = ({ value }) => {
    return new docx.Math({
        children: parseLatex(value).flatMap((runs) => runs),
    });
};
const buildText = (text, deco) => {
    return new docx.TextRun({
        text,
        bold: deco.strong,
        italics: deco.emphasis,
        strike: deco.delete,
    });
};
const buildBreak = (_) => {
    return new docx.TextRun({ text: "", break: 1 });
};
const buildLink = ({ children, url, title: _title }, ctx) => {
    const { nodes } = convertNodes(children, ctx);
    return new docx.ExternalHyperlink({
        link: url,
        children: nodes,
    });
};
const buildImage = ({ url, title: _title, alt: _alt }, images) => {
    const img = images[url];
    invariant(img, `Fetch image was failed: ${url}`);
    const { image, width, height } = img;
    return new docx.ImageRun({
        data: image,
        transformation: {
            width,
            height,
        },
    });
};
const buildFootnote = ({ children }, ctx) => {
    // FIXME: transform to paragraph for now
    const { nodes } = convertNodes(children, ctx);
    return new docx.Paragraph({
        children: nodes,
    });
};
const buildFootnoteDefinition = ({ children }, ctx) => {
    return {
        children: children.map((node) => {
            const { nodes } = convertNodes([node], ctx);
            return nodes[0];
        }),
    };
};
const buildFootnoteReference = ({ identifier }) => {
    // do we need Context?
    return new docx.FootnoteReferenceRun(parseInt(identifier));
};

const plugin = function (opts = {}) {
    let images = {};
    this.Compiler = (node) => {
        return mdastToDocx(node, opts, images);
    };
    return async (node) => {
        const imageList = [];
        unistUtilVisit.visit(node, "image", (node) => {
            imageList.push(node);
        });
        if (imageList.length === 0) {
            return node;
        }
        const imageResolver = opts.imageResolver;
        invariant(imageResolver, "options.imageResolver is not defined.");
        const imageDatas = await Promise.all(imageList.map(({ url }) => imageResolver(url)));
        images = imageList.reduce((acc, img, i) => {
            acc[img.url] = imageDatas[i];
            return acc;
        }, {});
        return node;
    };
};

Object.defineProperty(exports, "TableWidthType", {
    enumerable: true,
    get: function () { return docx.WidthType; }
});
exports.default = plugin;
//# sourceMappingURL=index.js.map
