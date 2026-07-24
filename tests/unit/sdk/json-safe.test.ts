import { describe, expect, it } from 'vitest';
import { JsonSnapshotLimitError, snapshotJsonValue } from '../../../src/sdk/adapters/json-safe';

describe('strict JSON snapshot budgets', () => {
  it.each([
    {
      value: [null, null],
      limits: { maxNodes: 2 },
      resource: 'nodes',
      actual: 3,
      limit: 2,
    },
    {
      value: { first: null, second: null },
      limits: { maxProperties: 1 },
      resource: 'properties',
      actual: 2,
      limit: 1,
    },
    {
      value: '€€',
      limits: { maxStringBytes: 5 },
      resource: 'stringBytes',
      actual: 6,
      limit: 5,
    },
    {
      value: [null, null],
      limits: { maxEstimatedBytes: 10 },
      resource: 'estimatedBytes',
      actual: 11,
      limit: 10,
    },
  ])(
    'fails closed with stable $resource diagnostics before completing an untrusted snapshot',
    ({ value, limits, resource, actual, limit }) => {
      expect(() => snapshotJsonValue(value, 'payload', limits)).toThrowError(
        expect.objectContaining({
          name: 'JsonSnapshotLimitError',
          message: `payload exceeds the maximum JSON snapshot ${resource}`,
          resource,
          actual,
          limit,
        }) satisfies Partial<JsonSnapshotLimitError>,
      );
    },
  );

  it('rejects giant shallow containers before reading their data properties', () => {
    let arrayDescriptorReads = 0;
    const array = new Proxy(
      Array.from({ length: 100 }, () => null),
      {
        getOwnPropertyDescriptor(target, property) {
          arrayDescriptorReads += 1;
          return Reflect.getOwnPropertyDescriptor(target, property);
        },
      },
    );
    let objectDescriptorReads = 0;
    const object = new Proxy(
      Object.fromEntries(Array.from({ length: 100 }, (_, index) => [`key-${index}`, null])),
      {
        getOwnPropertyDescriptor(target, property) {
          objectDescriptorReads += 1;
          return Reflect.getOwnPropertyDescriptor(target, property);
        },
      },
    );

    expect(() =>
      snapshotJsonValue(array, 'array', {
        maxNodes: 10,
        maxProperties: 10,
        maxEstimatedBytes: 100,
      }),
    ).toThrowError(JsonSnapshotLimitError);
    expect(() =>
      snapshotJsonValue(object, 'object', {
        maxNodes: 10,
        maxProperties: 10,
        maxEstimatedBytes: 100,
      }),
    ).toThrowError(JsonSnapshotLimitError);
    expect(arrayDescriptorReads).toBe(0);
    expect(objectDescriptorReads).toBe(0);
  });
});
