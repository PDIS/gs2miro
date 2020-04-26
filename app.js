const Koa = require('koa');
const Router = require('koa-router');
const views = require('koa-views');
const koaBody = require('koa-body');
const fs = require('fs')
const middlewares = require('./middlewares')

const app = new Koa();
const router = new Router();

app.use(koaBody());
app.use(views(`${__dirname}/views`, {
  extension: 'pug'
}));

router.get('/', async (ctx) => {
  await ctx.render('index', {
    title: 'Koa2',
    name: 'Rorast',
    engine: 'pug'
  })
});

router.post('/', async (ctx) => {
  const filename = `${ctx.request.body.board}.rtb`
  await middlewares.gs2mrio(ctx.request.body.sheet, ctx.request.body.board)
  ctx.body = await fs.createReadStream(filename);
  await ctx.attachment(filename)
  await middlewares.deletefile(filename)
  //https://docs.google.com/spreadsheets/d/e/2PACX-1vRTTodYxT8D5r_PS6GjHAoTMMFj-JQKHQ6o_GN7Srg_NvAJMa27ah1xtB5Z7Y2c0FrDGwddZTq0wgDs/pub?output=xlsx
});

app.use(router.routes());
app.listen(3000);