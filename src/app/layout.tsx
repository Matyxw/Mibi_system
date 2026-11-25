import type { Metadata } from 'next';
import { Plus_Jakarta_Sans } from 'next/font/google';
import { Analytics } from '@vercel/analytics/next';
import { ThemeProvider } from '@/components/theme-provider';
import './globals.css';

const plusJakarta = Plus_Jakarta_Sans({
  subsets: ['latin'],
  weight: ['400', '500', '600', '700'],
  display: 'swap'
});

const THEME_INIT_SCRIPT = `(function(){try{var d=localStorage.getItem('mibi_theme')==='dark';document.documentElement.classList.toggle('dark',d);document.documentElement.style.colorScheme=d?'dark':'light';}catch(e){}})();`;

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
    <html lang="es" suppressHydrationWarning>
      <head>
        <script dangerouslySetInnerHTML={{ __html: THEME_INIT_SCRIPT }} />
      </head>
      <body
        className={`${plusJakarta.className} bg-[#FAFAFA] text-[#111111] dark:bg-[#0B0B0D] dark:text-[#F4F4F5] antialiased selection:bg-[#FCE7F3] dark:selection:bg-[#F4B6CE]/25 dark:selection:text-[#FCE7F3] transition-colors duration-150`}
      >
        <ThemeProvider>{children}</ThemeProvider>
        <Analytics />
      </body>
    </html>
  );
}
