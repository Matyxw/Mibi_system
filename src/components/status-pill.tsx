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
      box: 'bg-emerald-50 text-emerald-700 border-emerald-200/60 dark:bg-emerald-500/[0.12] dark:text-emerald-300 dark:border-emerald-500/25',
      dot: 'bg-emerald-500 dark:bg-emerald-400'
    },
    amber: {
      box: 'bg-amber-50 text-amber-700 border-amber-200/60 dark:bg-amber-500/[0.12] dark:text-amber-300 dark:border-amber-500/25',
      dot: 'bg-amber-500 dark:bg-amber-400'
    },
    red: {
      box: 'bg-red-50 text-red-700 border-red-200/60 dark:bg-red-500/[0.12] dark:text-red-300 dark:border-red-500/25',
      dot: 'bg-red-500 dark:bg-red-400'
    },
    neutral: {
      box: 'bg-zinc-100 text-zinc-700 border-zinc-200/70 dark:bg-zinc-800/80 dark:text-zinc-300 dark:border-zinc-700/80',
      dot: 'bg-zinc-400 dark:bg-zinc-400'
    },
    ink: {
      box: 'bg-zinc-900/[0.06] text-zinc-900 border-zinc-300/80 dark:bg-zinc-100/10 dark:text-zinc-100 dark:border-zinc-500/40',
      dot: 'bg-zinc-900 dark:bg-zinc-100'
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
