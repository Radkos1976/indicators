var browserify = require('browserify-middleware');
var express = require('express');
const helmet = require('helmet');
const path = require('path');
const reload = require('reload');
const favicon = require('serve-favicon');
const cookieParser = require('cookie-parser');
const pg_settings=require('./settings/PostegresDB.js');
const FileStreamRotator = require('file-stream-rotator');
const logger = require('morgan');
const fs = require('fs');
const session = require('express-session');
const chk =require('./controlers/service_allow.js');
const Pers_ctrl= require('./routes/ctrl');
const bodyParser = require('body-parser');
const mainroutes = require('./routes/indicators');
const secure = require('./routes/secure');
const pg = require('pg');
const pgSession = require('connect-pg-simple')(session);
const compression = require('compression');
const minify = require('express-minify');
const uglifyEs = require('uglify-es');
var app = express();
var ifs_allow=true;
var pgPool = new pg.Pool(pg_settings.Session_manage);
pgPool.on('error', function (err, client) {
  console.error('idle client error', err.message, err.stack);
});
const logDirectory = path.join(__dirname, 'log');
fs.existsSync(logDirectory) || fs.mkdirSync(logDirectory);
// create a rotating write stream
const accessLogStream = FileStreamRotator.getStream({
  date_format: 'YYYYMMDD',
  filename: path.join(logDirectory, 'access-%DATE%.log'),
  frequency: 'daily',
  verbose: false
})
const execFile = require('child_process').execFile;
setTimeout(get_dataportal, 30000);
setTimeout(get_data, 60000);
CheckIsIFSallow();
// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'pug');
app.use(cookieParser());
app.use(session({
  store: new pgSession ({pool: pgPool}),
  unset: 'destroy',
  ttl: 30,
  saveUninitialized: true,
  secret: 'sess_uuid',
  resave: true,
  cookie: {maxAge: 30 * 60 * 1000, httpOnly:true} // 1 hour
}));
// uncomment after placing your favicon in /public
app.use(helmet());
//app.use(helmet.noCache());
app.use(logger(':date[iso] :method :remote-addr :url :status :response-time ', {stream: accessLogStream} , {flags: 'a'}));
app.use(compression());
app.use(minify({
  uglifyJsModule: uglifyEs
}));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
app.use('/ctrl_info',Pers_ctrl);
app.use('/secure',secure);
app.use('/bootstrap', express.static(__dirname + '/node_modules/bootstrap'));
app.use('/jquery', express.static(__dirname + '/node_modules/jquery/dist')); //jquery.js
app.use('/jquery-ui', express.static(__dirname + '/node_modules/jquery-ui-dist')); //jquery-ui.js ;jquery-ui.css ;jquery-ui.theme.css
app.use('/chartNew',express.static(__dirname + '/node_modules/chartnew.js')); //ChartNew.js
app.use('/ko', express.static(__dirname + '/node_modules/knockout/build/output'));

app.use(express.static(path.join(__dirname, 'public'),{  maxAge: 5000000 }));
app.use(express.static(path.join(__dirname, 'dist/boardinfo')));
app.use('/', mainroutes);
app.use(express.static(path.join(__dirname, 'dist/boardinfo')));
 //knockout-latest.js
app.use('/js', browserify(__dirname + '/node_modules/secure-remote-password', { standalone: 'srp',}));
app.use('/cjs', browserify(__dirname + '/node_modules/crypto-json', { standalone: 'cjson',}));

const reloadServer = reload(app,{verbose: true});
app.get('/refr', function (req, res) {
  reloadServer.reload()
  res.send('Service refresh DONE')
})
// catch 404 and forward to error handler
app.use(function(req, res, next) {
  var err = new Error('Not Found');
  err.status = 404;
  next(err);
});
async function CheckIsIFSallow() {
  try {
    ifs_allow= await chk.Chk_is_allow();
    console.log(ifs_allow);
  } catch (e) {
    console.log(e);
  } finally {
    setTimeout(CheckIsIFSallow, 200000);
  }
}
function get_dataportal() {
  try {
    if (ifs_allow==false)
    {
      const child = execFile('cscript.exe', ['run1.vbs'], (error, stdout, stderr) => {
      console.log(stdout);
      });
    };
  } catch (e) {
    console.log(e);
  } finally {
    setTimeout(get_dataportal, 3000000);
  }
}
function get_data(){
  try {
    if (ifs_allow==false)
    {
      const child1 = execFile('cscript.exe', ['run.vbs'], (error, stdout, stderr) => {
      console.log(stdout);
        });
    };
  } catch (e) {
    console.log(e);
  } finally {
    setTimeout(get_data, 600000);
  }
}
// error handlers

// development error handler
// will print stacktrace
if (app.get('env') === 'development') {
  app.use(function(err, req, res, next) {
    res.status(err.status || 500);
    res.render('error', {
      message: err.message,
      error: err
    });
  });
}
// production error handler
// no stacktraces leaked to user
app.use(function(err, req, res, next) {
  res.status(err.status || 500);
  res.render('error', {
    message: err.message,
    error: {}
  });
});
module.exports = app;
