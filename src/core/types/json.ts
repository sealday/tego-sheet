/** @inline */
export type JsonPrimitive = string | number | boolean | null;

/** @inline */
export type JsonArray = readonly JsonValue[];

/** @inline */
export interface JsonObject {
  /** A JSON property whose value is itself JSON-compatible. */
  readonly [key: string]: JsonValue;
}

/** A value that can be represented losslessly in JSON. */
export type JsonValue = JsonPrimitive | JsonArray | JsonObject;

/**
 * Preserves known property types while allowing additional JSON-compatible keys.
 *
 * @inline
 */
export type JsonExtensible<Known extends object> = Readonly<Known> & JsonObject;

/**
 * Represents a JSON object whose decimal keys form a sparse collection.
 * Missing indexes read as `undefined` and are omitted when serialized.
 *
 * @inline
 */
export type SparseJsonCollection<Known extends object = object> = Readonly<Known> & {
  /** JSON-compatible entry stored at a sparse decimal index. */
  readonly [key: string]: JsonValue | undefined;
};
