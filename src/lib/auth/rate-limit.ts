import 'server-only';

interface RateLimitBucket {
  count: number;
  resetAt: number;
}

const MAX_ATTEMPTS = 5;
const WINDOW_MS = 15 * 60 * 1000; // 15 minutos

const globalForRl = globalThis as unknown as {
  __mibiLoginRateLimit?: Map<string, RateLimitBucket>;
};

const getStore = (): Map<string, RateLimitBucket> => {
  if (!globalForRl.__mibiLoginRateLimit) {
    globalForRl.__mibiLoginRateLimit = new Map<string, RateLimitBucket>();
  }
  return globalForRl.__mibiLoginRateLimit;
};

export const checkLoginRateLimit = (ip: string, username: string): { allowed: boolean; retryAfterSec: number } => {
  const now = Date.now();
  const store = getStore();
  const keys = [`ip:${ip}`, `user:${username.toLowerCase()}`];

  for (const key of keys) {
    const bucket = store.get(key);
    if (bucket) {
      if (now >= bucket.resetAt) {
        store.delete(key);
      } else if (bucket.count >= MAX_ATTEMPTS) {
        return {
          allowed: false,
          retryAfterSec: Math.max(1, Math.ceil((bucket.resetAt - now) / 1000))
        };
      }
    }
  }
  return { allowed: true, retryAfterSec: 0 };
};

export const recordLoginFailure = (ip: string, username: string): void => {
  const now = Date.now();
  const store = getStore();
  const keys = [`ip:${ip}`, `user:${username.toLowerCase()}`];

  for (const key of keys) {
    const bucket = store.get(key);
    if (!bucket || now >= bucket.resetAt) {
      store.set(key, { count: 1, resetAt: now + WINDOW_MS });
    } else {
      bucket.count += 1;
      store.set(key, bucket);
    }
  }
};

export const clearLoginRateLimit = (ip: string, username: string): void => {
  const store = getStore();
  store.delete(`ip:${ip}`);
  store.delete(`user:${username.toLowerCase()}`);
};
