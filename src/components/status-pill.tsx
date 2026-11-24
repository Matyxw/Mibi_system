'use client';

import React from 'react';
import type { ToneStatus } from '@/lib/types';

export function StatusPill({
  tone = 'neutral',
  children
}: {
  tone?: ToneStatus;
  children: React.ReactNode;
}): React.JSX.Element {
  const styles: Record<ToneStatus, { box: string; dot: string }> = {
    green: {
      box: 'bg-emerald-50 text-emerald-700 border-emerald-200/60',
      dot: 'bg-emerald-500'
    },
    amber: {
      box: 'bg-amber-50 text-amber-700 border-amber-200/60',
      dot: 'bg-amber-500'
    },
    red: {
      box: 'bg-red-50 text-red-700 border-red-200/60',
      dot: 'bg-red-500'
    },
    neutral: {
      box: 'bg-zinc-100 text-zinc-700 border-zinc-200/70',
      dot: 'bg-zinc-400'
    },
    ink: {
      box: 'bg-zinc-900/[0.06] text-zinc-900 border-zinc-300/80',
      dot: 'bg-zinc-900'
    }
  };
  const s = styles[tone];
  return (
    <span
      className={`inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-medium border whitespace-nowrap ${s.box}`}
    >
      <span className={`w-1.5 h-1.5 rounded-full ${s.dot}`} />
      {children}
    </span>
  );
}
