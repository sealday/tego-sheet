import { describe, expect, it, vi } from 'vitest';
import { inputBytes } from '../../../src/interchange/contracts';

describe('interchange input allocation limits', () => {
  it('rejects an oversized Blob before materializing its ArrayBuffer', async () => {
    const arrayBuffer = vi.fn(async () => new ArrayBuffer(0));
    const input = {
      size: 1024,
      arrayBuffer,
    } as unknown as Blob;

    await expect(inputBytes(input, undefined, 16)).rejects.toThrow(/byte limit/u);
    expect(arrayBuffer).not.toHaveBeenCalled();
  });
});
