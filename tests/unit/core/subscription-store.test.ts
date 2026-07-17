import { expect, it } from 'vitest';
import { SubscriptionStore } from '../../../src/core/controller/subscription-store';

it('notifies the listener snapshot in subscription order', () => {
  const store = new SubscriptionStore<string>();
  const notifications: string[] = [];
  store.subscribe((value) => notifications.push(`first:${value}`));
  store.subscribe((value) => notifications.push(`second:${value}`));

  store.publish('ready');

  expect(notifications).toEqual(['first:ready', 'second:ready']);
});
