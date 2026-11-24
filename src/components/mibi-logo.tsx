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
  const src =
    variant === 'pink'
      ? '/mibi-logo-pink.png'
      : variant === 'white'
        ? '/mibi-logo-white.png'
        : '/mibi-logo-ink.png';

  return (
    <img
      src={src}
      alt="Mibi®"
      width={width}
      height={Math.round(width * (210 / 480))}
      style={{ width, height: 'auto' }}
      className="block select-none object-contain"
      draggable={false}
    />
  );
}
