# Smart Fix — Intelligent Problem Resolution

Analyze the issue and automatically select the best approach to fix it.

## Trigger
Use when the user reports a bug, error, or problem that needs diagnosis and fixing.

## Instructions

First, analyze the issue to categorize it, then apply the appropriate strategy:

### For Code Errors and Bugs
If the issue involves application errors, exceptions, or functional bugs:
1. Reproduce the error — identify exact steps and error messages
2. Trace the root cause through the call stack
3. Identify the minimal fix
4. Verify the fix doesn't introduce regressions
5. If complex, delegate to `/systematic-debugging` skill

### For Performance Issues
If the issue involves slow response times, high resource usage, or timeouts:
1. Profile the bottleneck (N+1 queries, blocking I/O, large payloads)
2. Measure before/after with concrete metrics
3. Apply targeted optimization
4. Verify no functionality regression

### For Deployment/Infrastructure Issues
If the issue involves deployment failures, environment config, or CI/CD:
1. Check environment variables and secrets
2. Review build logs for errors
3. Verify infrastructure configuration
4. Test locally vs production differences

### For Database Issues
If the issue involves slow queries, data integrity, or migration problems:
1. Analyze query execution plans
2. Check indexes and schema
3. Review connection pooling settings
4. Validate migration scripts

### For UI/Frontend Issues
If the issue involves rendering, state management, or user interaction:
1. Check browser console for errors
2. Verify component state and props
3. Test across browsers/devices
4. Check CSS specificity and layout

### Multi-Domain Issues
For complex issues spanning multiple areas:
1. Address the primary symptom first
2. Fix secondary issues in order of impact
3. Verify integration between fixes
4. Run full test suite

## Output
1. **Diagnosis**: What's wrong and why
2. **Root cause**: The underlying issue
3. **Fix**: The specific changes needed
4. **Verification**: How to confirm it's fixed
