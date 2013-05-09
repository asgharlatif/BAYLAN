(function(){

    var Vehicle= Backbone.Model.extend({
        defaults:
        {

            'type':'car',
            'color':'red'
        }

    });

    var car=new Vehicle();



//    var volcano= _.extend({},Backbone.Events);

    //volcano.on('disaster:eruption',function(option){
    //    console.log("eruption has started at event of disaster "+option.plan);
    //});

    //volcano.trigger('disaster:eruption',{plan:'run'});
    //volcano.trigger('disaster:eruption',{plan:'run 2'});
    //volcano.off('disaster:eruption');
    ///volcano.trigger('disaster:eruption',{plan:'run 3'});

    //console.log(ford.cid);
    //console.log(ford.isNew());
    console.log(car.get('color'));


})();