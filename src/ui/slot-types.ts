import type { ReactNode } from 'react';
import type { SheetTabsRenderProps, ToolbarRenderProps } from '../core';

/** Renders custom toolbar chrome from {@link ToolbarRenderProps} in place of the default toolbar. */
export type ToolbarRenderer = (props: ToolbarRenderProps) => ReactNode;

/** Renders custom tab chrome from {@link SheetTabsRenderProps} in place of the default tab bar. */
export type SheetTabsRenderer = (props: SheetTabsRenderProps) => ReactNode;
