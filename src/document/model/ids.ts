declare const documentIdBrand: unique symbol;
declare const documentSheetIdBrand: unique symbol;
declare const styleIdBrand: unique symbol;
declare const validationIdBrand: unique symbol;
declare const templateIdBrand: unique symbol;
declare const resourceIdBrand: unique symbol;
declare const bindingIdBrand: unique symbol;
declare const objectIdBrand: unique symbol;

export type DocumentId = string & { readonly [documentIdBrand]: true };
export type DocumentSheetId = string & { readonly [documentSheetIdBrand]: true };
export type StyleId = string & { readonly [styleIdBrand]: true };
export type ValidationId = string & { readonly [validationIdBrand]: true };
export type TemplateId = string & { readonly [templateIdBrand]: true };
export type ResourceId = string & { readonly [resourceIdBrand]: true };
export type BindingId = string & { readonly [bindingIdBrand]: true };
export type ObjectId = string & { readonly [objectIdBrand]: true };
