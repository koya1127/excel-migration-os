# Frontend Patterns & Performance (Next.js & TypeScript)

## 1. Next.js App Router & RSC
- **Server Components (RSC)**: Favor RSC for data fetching (default in Next.js 13+). Ensure `use client` is only at the necessary leaf node.
- **`use client` Overuse**: Flag `use client` on purely informational or layout components that don't need interactivity.
- **Server Actions**: Ensure Server Actions have proper CSRF protection (Next.js does some, but check manual logic) and input validation (e.g., Zod).

## 2. Performance & Rendering
- **Re-rendering**: Check for stable references in `useEffect` dependencies. Suggest `useMemo`/`useCallback` for expensive computations or callback props to memoized children.
- **Images**: Use `next/image` with proper `width`/`height` or `fill` to avoid Layout Shift (CLS).
- **Client Bundles**: Flag large external library imports. Suggest `dynamic(() => ...)` for heavy client-side components.

## 3. TypeScript & Data Integrity
- **Zod / Validation**: Prefer schema-based validation (Zod) for API responses and form inputs.
- **Non-Null Assertions (`!`)**: Discourage `!` unless strictly proven. Suggest optional chaining or proper null checks.
- **`any` Type**: Flag `any`. Suggest `unknown` or specific interfaces.

## 4. State Management
- **Local vs. Global**: Prefer React State/Context for UI state. Avoid Redux/Zustand unless global sync is truly required.
- **Prop Drilling**: Identify deep prop drilling and suggest Context or Composition.
- **Form State**: Recommend React Hook Form for complex forms to avoid global re-renders on every keystroke.
