const config = require('config')

module.exports.getTargetPathFile = () => {
  return config.get('App.file.target')
}
