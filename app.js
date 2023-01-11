const express = require('express')
const app = express()
const port = 3000

app.use(function(req, res, next){
   res.header("Access-Control-Allow-Origin", "*");
   res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Context-Type, Access");
   next();
});

app.get('/', (req, res) => {
    res.send({'text': 'this text was sent from server'})
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}!`);
});