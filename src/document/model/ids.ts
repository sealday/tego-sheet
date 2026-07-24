declare const documentIdBrand: unique symbol;
declare const documentSheetIdBrand: unique symbol;
declare const styleIdBrand: unique symbol;
declare const validationIdBrand: unique symbol;
declare const templateIdBrand: unique symbol;
declare const resourceIdBrand: unique symbol;
declare const bindingIdBrand: unique symbol;
declare const objectIdBrand: unique symbol;
declare const groupIdBrand: unique symbol;

/** Stable opaque identifier for a spreadsheet document. */
export type DocumentId = string & {
  /** Type-only brand that distinguishes document identifiers. */
  readonly [documentIdBrand]: true;
};
/** Stable opaque identifier for a sheet within a document. */
export type DocumentSheetId = string & {
  /** Type-only brand that distinguishes document sheet identifiers. */
  readonly [documentSheetIdBrand]: true;
};
/** Stable opaque identifier for a style registry entry. */
export type StyleId = string & {
  /** Type-only brand that distinguishes style identifiers. */
  readonly [styleIdBrand]: true;
};
/** Stable opaque identifier for a validation registry entry. */
export type ValidationId = string & {
  /** Type-only brand that distinguishes validation identifiers. */
  readonly [validationIdBrand]: true;
};
/** Stable opaque identifier for a spreadsheet template. */
export type TemplateId = string & {
  /** Type-only brand that distinguishes template identifiers. */
  readonly [templateIdBrand]: true;
};
/** Stable opaque identifier for a document resource. */
export type ResourceId = string & {
  /** Type-only brand that distinguishes resource identifiers. */
  readonly [resourceIdBrand]: true;
};
/** Stable opaque identifier for a data binding. */
export type BindingId = string & {
  /** Type-only brand that distinguishes binding identifiers. */
  readonly [bindingIdBrand]: true;
};
/** Stable opaque identifier for a document object. */
export type ObjectId = string & {
  /** Type-only brand that distinguishes object identifiers. */
  readonly [objectIdBrand]: true;
};
/** Stable opaque identifier for a worksheet outline group. */
export type GroupId = string & {
  /** Type-only brand that distinguishes outline group identifiers. */
  readonly [groupIdBrand]: true;
};
