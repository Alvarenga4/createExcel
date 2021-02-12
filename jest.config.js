const { resolve } = require('path')
const root = resolve(__dirname)

module.exports = {
  rootDir: root,
  displayName: 'root-tests',
  testEnvironment: 'node',
  clearMocks: true,
  moduleNameMapper: {
    '@test/(.*)': '<rootDir>/test/$1'
  }
}
