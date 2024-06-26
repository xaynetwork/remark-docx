import { Paragraph, ParagraphChild, Table, TableOfContents, IImageOptions, ITableWidthProperties } from "docx";
import type { IPropertiesOptions } from "docx/build/file/core-properties";
import type * as mdast from "./models/mdast";
export type ImageDataMap = {
    [url: string]: ImageData;
};
export type ImageData = {
    image: IImageOptions["data"];
    width: number;
    height: number;
};
export type ImageResolver = (url: string) => Promise<ImageData> | ImageData;
type customStyles = Readonly<{
    table?: {
        width?: ITableWidthProperties | undefined;
    };
}>;
export interface DocxOptions extends Pick<IPropertiesOptions, "title" | "subject" | "creator" | "keywords" | "description" | "lastModifiedBy" | "revision" | "styles" | "background"> {
    /**
     * Set output type of `VFile.result`. `buffer` is `Promise<Buffer>`. `blob` is `Promise<Blob>`.
     */
    output?: "buffer" | "blob";
    /**
     * **You must set** if your markdown includes images. See example for [browser](https://github.com/inokawa/remark-docx/blob/main/stories/playground.stories.tsx) and [Node.js](https://github.com/inokawa/remark-docx/blob/main/src/index.spec.ts).
     */
    imageResolver?: ImageResolver;
    customStyles?: customStyles;
}
type DocxChild = Paragraph | Table | TableOfContents;
type DocxContent = DocxChild | ParagraphChild;
export interface Footnotes {
    [key: string]: {
        children: Paragraph[];
    };
}
export interface ConvertNodesReturn {
    nodes: DocxContent[];
    footnotes: Footnotes;
}
export declare const mdastToDocx: (node: mdast.Root, { output, title, subject, creator, keywords, description, lastModifiedBy, revision, styles, background, customStyles, }: DocxOptions, images: ImageDataMap) => Promise<any>;
export {};
