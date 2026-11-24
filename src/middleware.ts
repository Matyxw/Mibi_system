import { NextResponse, type NextRequest } from 'next/server';
import {
  getSessionCookieName,
  SESSION_DURATION_SECONDS,
  signSessionToken,
  verifySessionToken
} from '@/lib/auth/jwt';

const PUBLIC_PATHS = new Set(['/login', '/robots.txt']);

export async function middleware(request: NextRequest): Promise<NextResponse> {
  const { pathname } = request.nextUrl;

  if (
    pathname.startsWith('/_next/') ||
    pathname === '/favicon.ico' ||
    pathname === '/favicon.png' ||
    pathname === '/favicon.svg' ||
    pathname === '/mibi-logo-pink.png' ||
    pathname === '/mibi-logo-ink.png' ||
    pathname === '/mibi-logo-white.png'
  ) {
    return NextResponse.next();
  }

  const cookieName = getSessionCookieName();
  const token = request.cookies.get(cookieName)?.value ?? '';
  const session = await verifySessionToken(token);

  if (PUBLIC_PATHS.has(pathname)) {
    if (pathname === '/login' && session) {
      return NextResponse.redirect(new URL('/', request.url));
    }
    return NextResponse.next();
  }

  if (!session) {
    if (pathname.startsWith('/api/')) {
      return NextResponse.json({ error: 'No autorizado' }, { status: 401 });
    }
    const loginUrl = new URL('/login', request.url);
    return NextResponse.redirect(loginUrl);
  }

  // Expiración deslizante de 12h en cada navegación autenticada
  const response = NextResponse.next();
  const refreshedToken = await signSessionToken(session);
  response.cookies.set(cookieName, refreshedToken, {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax',
    path: '/',
    maxAge: SESSION_DURATION_SECONDS
  });
  return response;
}

export const config = {
  matcher: ['/((?!_next/static|_next/image).*)']
};
