import type { NextConfig } from "next";

// NEXT_PUBLIC_API_URL is required in production for rewrites and CSP.
// Vercel sets it at build time via environment variables dashboard.
// Falls back to localhost for local development only.
const resolvedBackendUrl = process.env.NEXT_PUBLIC_API_URL || "http://localhost:5269";

const nextConfig: NextConfig = {
  // Proxy JSON-only backend API calls through Next.js rewrites.
  // File upload endpoints (scan, extract, upload, migrate) call the backend directly
  // via NEXT_PUBLIC_API_URL to avoid Vercel's 4.5MB body size limit.
  async rewrites() {
    return [
      { source: "/api/convert", destination: `${resolvedBackendUrl}/api/convert` },
      { source: "/api/convert/:path*", destination: `${resolvedBackendUrl}/api/convert/:path*` },
      { source: "/api/deploy", destination: `${resolvedBackendUrl}/api/deploy` },
    ];
  },
  async headers() {
    return [
      {
        source: "/(.*)",
        headers: [
          {
            key: "Content-Security-Policy",
            value: [
              "default-src 'self'",
              // Clerk needs inline scripts/styles and its own domain
              "script-src 'self' 'unsafe-inline' https://*.clerk.accounts.dev https://challenges.cloudflare.com",
              "style-src 'self' 'unsafe-inline'",
              "img-src 'self' data: https://*.clerk.com https://img.clerk.com",
              // Connect: backend API, Clerk, Stripe, Google APIs
              `connect-src 'self' ${resolvedBackendUrl} https://*.clerk.accounts.dev https://api.stripe.com https://accounts.google.com https://oauth2.googleapis.com`,
              "frame-src 'self' https://*.clerk.accounts.dev https://challenges.cloudflare.com https://js.stripe.com",
              "font-src 'self'",
              "object-src 'none'",
              "base-uri 'self'",
              "form-action 'self'",
            ].join("; "),
          },
          {
            key: "X-Content-Type-Options",
            value: "nosniff",
          },
          {
            key: "Referrer-Policy",
            value: "strict-origin-when-cross-origin",
          },
          {
            key: "X-Frame-Options",
            value: "DENY",
          },
          {
            key: "Strict-Transport-Security",
            value: "max-age=31536000; includeSubDomains",
          },
          {
            key: "Permissions-Policy",
            value: "camera=(), microphone=(), geolocation=()",
          },
        ],
      },
    ];
  },
};

export default nextConfig;
