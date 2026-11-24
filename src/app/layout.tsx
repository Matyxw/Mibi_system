import type { Metadata } from 'next';
import { Plus_Jakarta_Sans } from 'next/font/google';
import './globals.css';

const plusJakarta = Plus_Jakarta_Sans({
  subsets: ['latin'],
  weight: ['400', '500', '600', '700'],
  display: 'swap'
});

export const metadata: Metadata = {
  title: 'Mibi® — Gestión Mayorista',
  description: 'Sistema Operativo Mayorista Mibi®',
  robots: {
    index: false,
    follow: false,
    nocache: true
  },
  icons: {
    icon: [
      { url: '/favicon.ico', sizes: 'any' },
      { url: '/favicon.png', type: 'image/png', sizes: '64x64' },
      { url: '/favicon.svg', type: 'image/svg+xml' }
    ],
    shortcut: '/favicon.ico'
  }
};

export default function RootLayout({ children }: { children: React.ReactNode }): React.JSX.Element {
  return (
    <html lang="es">
      <body
        className={`${plusJakarta.className} bg-[#FAFAFA] text-[#111111] antialiased selection:bg-[#FCE7F3]`}
      >
        {children}
      </body>
    </html>
  );
}
