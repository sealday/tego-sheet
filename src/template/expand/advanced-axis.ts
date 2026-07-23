import type { TemplateIRBinding } from '../model';

export type StructuralAxis = 'vertical' | 'horizontal' | 'conditional';

export function structuralAxis(binding: TemplateIRBinding | undefined): StructuralAxis | undefined {
  if (binding?.type === 'repeat-rows') return 'vertical';
  if (binding?.type === 'repeat-columns') return 'horizontal';
  if (binding?.type === 'conditional-range') return 'conditional';
  if (binding?.type === 'repeat-range' && binding.axis !== 'both') return binding.axis;
  return undefined;
}
