const axios = require('axios')

const api = axios.create({
  baseURL: 'http://localhost:3001',
});

if (process.env.NODE_ENV === 'development') {
  api.get = () => {
    const data = require('./company.json')
    return { data };
  }
}

module.exports = api;