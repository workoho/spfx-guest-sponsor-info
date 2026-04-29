// ESLint configuration for the Azure Function (getGuestSponsors).
// root: true prevents this config from merging with the root SPFx config.
module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  parserOptions: {
    project: './tsconfig.json',
    tsconfigRootDir: __dirname,
    ecmaVersion: 2020,
    sourceType: 'module',
  },
  plugins: ['@typescript-eslint', 'promise'],
  extends: [
    'plugin:@typescript-eslint/recommended',
    'plugin:promise/recommended',
  ],
  rules: {
    // Floating promises are bugs in async code — keep as error.
    '@typescript-eslint/no-floating-promises': 'error',
    // Explicit return types aid readability; allow expression inference.
    '@typescript-eslint/explicit-function-return-type': [
      'warn',
      { allowExpressions: true, allowTypedFunctionExpressions: true },
    ],
  },
};
