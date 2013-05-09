(function(){

    var Vehicle=Backbone.Model.extend({})
    var Car=new Vehicle({
        'Color':'Red'
    });

    var WorkSpace=Backbone.Router.extend({

        routes:{
            'book/:query':'Book',
            'search/:query':'search',
            'ListData/:DataType':'DataListening'

        },

        Book:function(query){
            $('body').html("Searched for book page");
        },
        search:function(query){
            //console.log();
            $('body').html("Searched for" + query);
        },
        DataListening:function(DataType){
            //console.log();
            $('body').html("Data Listed is of Type "+ DataType);
        }

    });

    var Router=new WorkSpace();

    Backbone.history.start();

    Router.navigate('book/asghar',{trigger:true});

})();