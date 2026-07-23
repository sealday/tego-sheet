import type { TemplateIRBinding } from '../model';

export type StructuralAxis = 'vertical' | 'horizontal' | 'both' | 'conditional';

export function structuralAxis(binding: TemplateIRBinding | undefined): StructuralAxis | undefined {
  if (binding?.type === 'repeat-rows') return 'vertical';
  if (binding?.type === 'repeat-columns') return 'horizontal';
  if (binding?.type === 'conditional-range') return 'conditional';
  if (binding?.type === 'repeat-range') return binding.axis;
  return undefined;
}
