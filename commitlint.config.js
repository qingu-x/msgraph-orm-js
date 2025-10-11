export default { 
  extends: ['@commitlint/config-conventional'],
  rules: {
    'type-enum': [2, 'always', ['feat', 'fix', 'docs', 'chore', 'style', 'refactor', 'test', 'build', 'ci', 'perf', 'revert', 'docs', 'chore', 'style', 'refactor', 'test', 'build', 'ci', 'perf', 'revert']]
  }
};
