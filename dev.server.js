'use strict';
const colors = require('colors');
const webpack = require('webpack');
const WebpackDevServer = require('webpack-dev-server');
const fs = require('fs');

let webpackConfig = require('./webpack.config.js');
webpackConfig.entry.unshift('webpack-dev-server/client?https://localhost:8080/');
const compiler = webpack(webpackConfig);

let server = new WebpackDevServer(compiler, {
  contentBase: 'bin/',
  publicPath: '/js/',
  compress: true,
  https: true,
  proxy: {'/rest': 'http://localhost:8081'}
});
server.listen(8080, 'localhost', function () {
});

let Filter = require('./src/Filter');

let express = require('express');
let app = express();
let http = require('http');
let https = require('https');
let request = require('request');
let bodyParser = require('body-parser');
app.use( bodyParser.json() );       // to support JSON-encoded bodies
app.use(bodyParser.urlencoded({     // to support URL-encoded bodies
  extended: true
}));

app.get('/', function (req, res) {
  res.send('Hello World!');
});

app.get('/rest', function (req, res) {
  res.send('Hello World! (rest)');
});

app.get('/rest/search', function (req, res) {
  res.send('Hello World! (rest-serach)');
});

app.post('/rest/search/v2', function (req, res) {

  let params = {
    '$top': 25,
    '$orderby': 'ReceivedDateTime desc'
  };

  let options = {
    url: 'https://outlook.office.com/api/v2.0/me/mailfolders/AllItems/messages?',
    headers: {
      'Content-Type': 'application/json;charset=UTF-8',
      'Authorization': req.headers.authorization,
      'X-AnchorMailbox': 'jdevost@coveo.com'
    }
  };

  if (req.body.aq) {
    console.log('AQ: ', req.body.aq);
    let filter = new Filter(req.body.aq);
    params['$filter'] = filter.generateOutlookFilter();
    params['$orderby'] = filter.getFields().join() + ',' + params['$orderby'];
  }

  if (req.body.q) {
    params['$search'] = '"' + req.body.q + '"';
    // can't have $orderBy with $search
    delete params['$orderby'];
    delete params['$filter'];
  }

  let aParams = [];
  for (var p in params) {
    aParams.push( [p, params[p]].join('=') );
  }
  options.url += aParams.join('&');

  console.log('Request = ', options.url);

  let callback = (error, response, body) => {
    let json = {error: 'invalid request'};
    if (!error && response.statusCode === 200) {
      let outlookJson = JSON.parse(body);

      json = {
        totalCount : outlookJson.value.length,
        termsToHighlight: {},
        phrasesToHighlight: {},
        queryCorrections: [],
        groupByResults: [],
        results: outlookJson.value.map( v =>{
          v.titleHighlights = [ ];
          v.firstSentencesHighlights = [ ];
          v.excerptHighlights = [ ];
          v.printableUriHighlights = [ ];
          v.summaryHighlights = [ ];
          v.parentResult = null;
          v.childResults = [ ];
          v.totalNumberOfChildResults = 0;

          v.clickUri = v.WebLink;
          v.excerpt = v.BodyPreview;
          v.title = v.Subject;
          v.raw = {
            objecttype: 'Email',
            date: v.ReceivedDateTime,
            filetype: 'email'
          };
          return v;
        })
      };

      // groupByResults for facets
      let facets = {}, fields = ['From', 'Importance'];
      outlookJson.value.forEach( v =>{
        fields.forEach( field=>{
          let value = v[field], caption = null;
          if (value && value.EmailAddress) {
            caption = value.EmailAddress.Name;
            value = value.EmailAddress.Address;
          }
          if (!value) {
            return;
          }
          let facet = facets[field];
          if (!facet) {
            facets[field] = {
              field: field,
              values: {}
            };
            facet = facets[field];
          }
          let valuesMap = facet.values[value];
          if (!valuesMap) {
            facet.values[value] = {
              value: value,
              lookupValue: caption || value,
              numberOfResults: 0
            };
            valuesMap = facet.values[value];
          }
          valuesMap.numberOfResults++;
        });
      });

      let aFacets = [];
      for (let f in facets) {
        let aValues = [];
        for (let v in facets[f].values) {
          aValues.push(facets[f].values[v]);
        }
        facets[f].values = aValues.sort((a,b)=>{
          return b.numberOfResults>a.numberOfResults?1:-1;
        });
        aFacets.push(facets[f]);
      }
      json.groupByResults = aFacets;
    }
    else {
      res.status(response.statusCode).type('application/json').send( JSON.stringify({message: response.statusMessage}) );
      return;
    }
    res.status(200).type('application/json').send( JSON.stringify(json) );
  };

  request(options, callback);

});

app.listen(8081, function () {
  console.log('Example app listening on port 8081!');
});
