import type { NextConfig } from "next";

// NEXT_PUBLIC_API_URL is required in production for rewrites and CSP.
// Vercel sets it at build time via environment variables dashboard.
// Falls back to localhost for local development only.
const resolvedBackendUrl = process.env.NEXT_PUBLIC_API_URL || "http://localhost:5269";

const nextConfig: NextConfig = {
  // All backend API calls go directly to NEXT_PUBLIC_API_URL from the browser.
  // No rewrites needed — avoids Vercel's 4.5MB body limit and 30s timeout.
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
