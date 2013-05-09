(function(){



  var Book=Backbone.Model.extend({
      url:'http://localhost:200/Backboneserver/backboneserver.asp'
  })

    var midnight=new Book({
        'title':'Midnight in the garden of good and eveil',
        'author':'Muhammad Asghar Ali'
    });

    midnight.save({},{

        success:function(){alert('done')},
        error:function(){alert('error')}

    });

    //$(alert('it is an alert message'));

})();