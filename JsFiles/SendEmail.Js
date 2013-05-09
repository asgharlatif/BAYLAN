
    (function () {


        var FormView = Backbone.View.extend({
            tagName: 'div',
            className: 'email-content',

            initialize: function () {              
            },

            events: {
                'click #Send-btn': 'handleSendBtnClick',
                'click #Send-Page': 'handleSendPageClick',
                'click #Send-Page-C': 'handleSendPageC',
                'keyup .search-query': 'nullcheckaction',

            },
            render: function () {                
                
                this.$el.fadeIn(5000);
                var Template = $('#Email-Template').html();
                this.$el.html(Template + '<span id="idreqemail" style="width:100px;color:red;display:None" >Required.</span><a id="Send-Page" href="#" class="btn btn-info ">Send</a>&nbsp;&nbsp;<a id="Send-Page-C" href="#" class="btn btn-info ">Cancel</a><br />');
                this.setDimensions(this.model.get('OpenDivHeight'));
                this.setColors('white');
                this.setBorders();
                
                return this;
            },
            handleSendBtnClick: function () {                               
                return this;
            },
            handleSendPageClick: function () {                
                var nullcheck = this.nullcheckaction();
                
                if (nullcheck == true) {
                    var emailaddress = $('#IdEmailId').val();
                    var successmessage = '<span class="centeralignlabeling" >'.concat(this.model.get('EmailSentSuccessMessage')).concat('</span>');

                    if (isValidEmailAddress(emailaddress)) {
                        var thisel = this.$el;

                        



                        this.SendEmail(thisel,emailaddress, function () {
                           
                            thisel.html(successmessage).fadeOut(5000, function () {
                                    $("#Send-btn").show();
                            });
                        });                       
                       
                    }
                    else {                       
                        $('#idreqemail').html("Invalid email");
                        $('#idreqemail').show(2000);
                    }
                    
                }
                return this;
            },
            setDimensions: function (SetHeight) {                
                this.$el.css({
                    height: SetHeight + 'px'
                });
            },
            setColors: function (BackColor) {
                this.$el.css({
                    background: BackColor
                });
            },
            setBorders: function () {
                this.$el.css({
                    border: 'thin dotted #808080'
                });
            },
            handleSendPageC: function () {
                this.$el.html('<span class="centeralignlabeling" >Thank you for viewing this feature. You may share it at later time.</span>').fadeOut(5000, function () {
                    $("#Send-btn").show();
                });
            },
            nullcheckaction: function () {
                var isnullcheckaction;
                $('#idreqemail').html("Required.");
                if (this.nullcheck()) {

                    $('#idreqemail').hide(2000);
                    isnullcheckaction = true;
                }
                else {
                    $('#idreqemail').show(2000);
                    isnullcheckaction = false;
                }

                return isnullcheckaction;

            },
            nullcheck: function () {
                var isnull;
                var idval = $('#IdEmailId').val();
                if (idval != "") {
                    isnull=true;
                }
                else {
                    isnull = false;
                }
                return isnull;
            },
            SendEmail: function (thisel,param, callback) {
                
                thisel.html('<span class="centeralignRedlabeling" >Please wait . It may take few seconds</span>');

                var blink = function () {
                    
                    $(thisel).fadeToggle("slow", "linear");
                };

                var timerId = 0;
                timerId=setInterval(blink, 1000);

                var productid = $('#txtproductid').val();
                //'description,itemcode

                var description=$('#txtdescription').val();
                var itemcode=$('#txtitemcode').val();



                SendDataToServer(timerId, productid,description,itemcode, param, function () {

                    clearInterval(timerId);
                    callback();

                });

                
            }

        });

        var datamodel = Backbone.Model.extend({});
        var dm = new datamodel({
            EmailText: ' Please enter user email address to share this page', Width: '300',
            EmailSentSuccessMessage: 'Page Send at your specified email address.',
            URLToRedirect:'http://localhost:200',
            //URLToRedirect:'http://baylan.com',

            PlainHeight: '25',
            OpenDivHeight:'31'

        });
        var V = new FormView({ model: dm });


        $(document).ready(function() {
            $("#Send-btn").click(function () {             
                $("#email-content-placing").append(V.render().el).show();               
                $("#Send-btn").hide();
            });            
        });

        function isValidEmailAddress(emailAddress) {
            var pattern = new RegExp(/^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?$/i);
            return pattern.test(emailAddress);
        };


        function SendDataToServer(timerId, productid,description,itemcode, param, callback) {



            var DriverURLPath = "http://localhost:200/driverdetail.asp?productid=" + productid;
            //var DriverURLPath = "http://baylan.com/driverdetail.asp?productid=" + productid;

            var DriverData={DriverList:[
                {ProductId:productid,ProductName:description,ProductModel:itemcode,DriverURL:DriverURLPath}
            ]};

            var renderedTemplate=Handlebars.templates.EmailTemplates(DriverData);
            var ParemString = "?ToEmail=" + param + "&pathname=" + renderedTemplate ;
           
            $.ajax
            (
            {
                type: "POST",
                url: "ServerSideOperation.asp" + ParemString,
                data: ParemString,
                success:function (msg) {
                //--------------------------2
               
                var list = $(msg).find('#IdEmailId').val();

                
                if (list == "Success") {
                                    
                    callback();
                }
                

                else if (list == "P") {
                    $(alert("Problem Sending email."));

                }                

                //--------------------------2
            }
            }
            );

            
        }



    })();
