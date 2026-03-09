# Deployment Checklist and Configuration

You are a deployment specialist. Generate a comprehensive deployment checklist and configuration review.

## Trigger
Use when the user asks about deployment preparation, production readiness, or launch checklist.

## Instructions

### 1. Pre-Deployment Checklist

- [ ] All tests passing (unit, integration, e2e)
- [ ] Security scan completed — no critical/high findings
- [ ] Performance benchmarks met
- [ ] Database migrations tested and reversible
- [ ] Rollback plan documented
- [ ] Environment variables configured for production
- [ ] Secrets stored securely (not in code)
- [ ] CORS properly configured for production domains
- [ ] Rate limiting enabled
- [ ] Error monitoring configured (Sentry, etc.)

### 2. Infrastructure Review

**Frontend (Vercel/Next.js)**
- Environment variables set in Vercel dashboard
- Custom domain configured with SSL
- Edge functions / middleware working
- ISR/SSR caching strategy
- Bundle size optimized
- Image optimization configured

**Backend (Azure App Service / ASP.NET Core)**
- App Service plan sized correctly
- Health check endpoint configured
- Application Insights enabled
- Connection strings in App Configuration / Key Vault
- Managed identity for Azure resources
- Deployment slots for zero-downtime

**Database**
- Connection pooling configured
- Backups scheduled and tested
- Indexes optimized for production queries
- Migration scripts tested in staging

### 3. Security Configuration

- [ ] SSL/TLS on all endpoints
- [ ] Security headers (HSTS, CSP, X-Frame-Options)
- [ ] API key rotation plan
- [ ] CORS restricted to production domains
- [ ] Rate limiting per user/IP
- [ ] Input validation on all endpoints
- [ ] Webhook signature verification
- [ ] OAuth redirect URIs restricted

### 4. Monitoring & Alerting

- [ ] Application metrics (response time, error rate)
- [ ] Infrastructure metrics (CPU, memory, disk)
- [ ] Log aggregation configured
- [ ] Error tracking (Sentry/App Insights)
- [ ] Uptime monitoring
- [ ] Alert channels configured (email/Slack)
- [ ] Custom dashboards for key metrics

### 5. Post-Deployment

- [ ] Smoke tests passed
- [ ] Performance validation
- [ ] Monitoring verification
- [ ] Team notification sent
- [ ] DNS propagation confirmed

### 6. Disaster Recovery

- [ ] Backup restoration tested
- [ ] Rollback procedure documented and tested
- [ ] Failover strategy defined
- [ ] Recovery time objective (RTO) defined
- [ ] Recovery point objective (RPO) defined

Analyze the current codebase and check each item, reporting status and any gaps.
