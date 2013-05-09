(function() {
  var template = Handlebars.template, templates = Handlebars.templates = Handlebars.templates || {};
templates['dblist'] = template(function (Handlebars,depth0,helpers,partials,data) {
  this.compilerInfo = [2,'>= 1.0.0-rc.3'];
helpers = helpers || Handlebars.helpers; data = data || {};
  var buffer = "", stack1, functionType="function", escapeExpression=this.escapeExpression, self=this;

function program1(depth0,data) {
  
  var buffer = "";
  buffer += "\r\n  <li><h1>"
    + escapeExpression((typeof depth0 === functionType ? depth0.apply(depth0) : depth0))
    + "</h1>\r\n  <h2> This is additional heading for db detail </h2>\r\n\r\n  </li>\r\n  ";
  return buffer;
  }

  buffer += "<ul class=\"db_list\">\r\n  ";
  stack1 = helpers.each.call(depth0, depth0.dbname, {hash:{},inverse:self.noop,fn:self.program(1, program1, data),data:data});
  if(stack1 || stack1 === 0) { buffer += stack1; }
  buffer += "\r\n</ul>";
  return buffer;
  });
templates['EmailTemplates'] = template(function (Handlebars,depth0,helpers,partials,data) {
  this.compilerInfo = [2,'>= 1.0.0-rc.3'];
helpers = helpers || Handlebars.helpers; data = data || {};
  var buffer = "", stack1, options, functionType="function", escapeExpression=this.escapeExpression, self=this, blockHelperMissing=helpers.blockHelperMissing;

function program1(depth0,data) {
  
  var buffer = "", stack1;
  buffer += "\r\n\r\n<ul>\r\n<li><b>Product Id:-</b>";
  if (stack1 = helpers.ProductId) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.ProductId; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + " </li>\r\n<li>Product Name:-";
  if (stack1 = helpers.ProductName) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.ProductName; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + " </li>\r\n<li>Product Model:-";
  if (stack1 = helpers.ProductModel) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.ProductModel; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + " </li>\r\n</ul>\r\n\r\n<br><br>\r\nDriver URL:- <br/>\r\n<a href= ";
  if (stack1 = helpers.DriverURL) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.DriverURL; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + ">";
  if (stack1 = helpers.DriverURL) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.DriverURL; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + "</a>\r\n\r\n<br><br>\r\n\r\n\r\n\r\n\r\n<br>\r\nSupport Team\r\n<br>\r\nBayLan / Taiwan.\r\n<br><br>\r\n\r\n\r\n";
  return buffer;
  }

  buffer += "<h3>\r\nDear User,\r\n</h3>\r\n\r\nClick over below link to download driver.\r\n<br><br>\r\nDriver detail is as below.\r\n<br><br>\r\n";
  options = {hash:{},inverse:self.noop,fn:self.program(1, program1, data),data:data};
  if (stack1 = helpers.DriverList) { stack1 = stack1.call(depth0, options); }
  else { stack1 = depth0.DriverList; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  if (!helpers.DriverList) { stack1 = blockHelperMissing.call(depth0, stack1, options); }
  if(stack1 || stack1 === 0) { buffer += stack1; }
  return buffer;
  });
templates['UserListening'] = template(function (Handlebars,depth0,helpers,partials,data) {
  this.compilerInfo = [2,'>= 1.0.0-rc.3'];
helpers = helpers || Handlebars.helpers; data = data || {};
  var buffer = "", stack1, options, functionType="function", escapeExpression=this.escapeExpression, self=this, blockHelperMissing=helpers.blockHelperMissing;

function program1(depth0,data) {
  
  var buffer = "", stack1;
  buffer += "\r\n        <tr>\r\n          <td>";
  if (stack1 = helpers.username) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.username; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + "</td>\r\n          <td>";
  if (stack1 = helpers.firstName) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.firstName; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + " ";
  if (stack1 = helpers.lastName) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.lastName; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + "</td>\r\n          <td>";
  if (stack1 = helpers.email) { stack1 = stack1.call(depth0, {hash:{},data:data}); }
  else { stack1 = depth0.email; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  buffer += escapeExpression(stack1)
    + "</td>\r\n        </tr>\r\n      ";
  return buffer;
  }

  buffer += " <table style=\"width:100%;background:red\" >\r\n    <thead>\r\n      <th>Username</th>\r\n      <th>Real Name</th>\r\n      <th>Email</th>\r\n    </thead>\r\n    <tbody>\r\n      ";
  options = {hash:{},inverse:self.noop,fn:self.program(1, program1, data),data:data};
  if (stack1 = helpers.users) { stack1 = stack1.call(depth0, options); }
  else { stack1 = depth0.users; stack1 = typeof stack1 === functionType ? stack1.apply(depth0) : stack1; }
  if (!helpers.users) { stack1 = blockHelperMissing.call(depth0, stack1, options); }
  if(stack1 || stack1 === 0) { buffer += stack1; }
  buffer += "\r\n    </tbody>\r\n  </table>";
  return buffer;
  });
})();