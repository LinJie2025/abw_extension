export default {
  extends: ['@commitlint/config-conventional'],
  rules: {
    'type-enum': [2, 'always', ['feat', 'fix', 'refactor', 'docs', 'chore', 'test', 'perf', 'style', 'build', 'ci', 'revert']],
    'subject-case': [0],
    'header-max-length': [2, 'always', 72]
  }
}
