'use client';

import React from 'react';

export type LogoVariant = 'ink' | 'pink' | 'white';

export function MibiLogo({
  width = 88,
  variant = 'ink'
}: {
  width?: number;
  variant?: LogoVariant;
}): React.JSX.Element {
  const height = Math.round(width * (210 / 480));

  if (variant === 'ink') {
    return (
      <>
        <img
          src="/mibi-logo-ink.png"
          alt="Mibi®"
          width={width}
          height={height}
          style={{ width, height: 'auto' }}
          className="block dark:hidden select-none object-contain"
          draggable={false}
        />
        <img
          src="/mibi-logo-white.png"
          alt="Mibi®"
          width={width}
          height={height}
          style={{ width, height: 'auto' }}
          className="hidden dark:block select-none object-contain"
          draggable={false}
        />
      </>
    );
  }

  const src = variant === 'pink' ? '/mibi-logo-pink.png' : '/mibi-logo-white.png';

  return (
    <img
      src={src}
      alt="Mibi®"
      width={width}
      height={height}
      style={{ width, height: 'auto' }}
      className="block select-none object-contain"
      draggable={false}
    />
  );
}
