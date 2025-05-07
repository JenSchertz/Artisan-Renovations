<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->
<%@ LANGUAGE=VBScript %>

<%

sHTML = "<HTML>"
sHTML = sHTML & "<HEAD><style>td {font-family:verdana,arial; font-size:8pt}</style></head><BODY>"
sHTML = sHTML & "<TABLE WIDTH=600 BORDER=0 CELLSPACING=2><TR><TD COLSPAN=2>"
sHTML = sHTML & "<P style=font-size:11pt;color:red><b>Request for Project Bid</b></P></TD></TR><TR><TD colspan=2>&nbsp;</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP VALIGN=TOP><B>Customer Name: </B></TD><TD WIDTH=450>" &  Request.Form("fname") & " " & Request.Form("lname") & "</TD></TR>"
sHTML = sHTML & "<TR><TD COLSPAN=2><B>---------------------------------------------------------------------------------------</B></TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Project Type: </B></TD><TD WIDTH=450>" & Request.Form("Project") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Design Status: </B></TD><TD WIDTH=450>" & Request.Form("Status") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Preferred Start Date: </B></TD><TD WIDTH=450>" & Request.Form("StartDate") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Customer's Timing is Flexible: </B></TD><TD WIDTH=450>" & Request.Form("TimingFlexible") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Working with Other Contractors: </B></TD><TD WIDTH=450>" & Request.Form("Others") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Customer Has Bid : </B></TD><TD WIDTH=450>" & Request.Form("HaveBid") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Project Description: </B></TD><TD WIDTH=450>" & Request.Form("description") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Project Address: </B></TD><TD WIDTH=450>" & Request.Form("address") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>City, State: </B></TD><TD WIDTH=450>" & Request.Form("city") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Project Zip Code: </B></TD><TD WIDTH=450>" & Request.Form("zip") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Phone Number:  </B></TD><TD WIDTH=450>" & Request.Form("phone") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & "<b>Alternate Phone: </b>" & Request.Form ("alternate") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Best Time to Call:  </B></TD><TD WIDTH=450>" & Request.Form("besttime") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Customer's E-Mail Address: </B></TD><TD WIDTH=450>" & Request.Form("email") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Customer's Fax Number: </B></TD><TD WIDTH=450>" &  Request.Form("fax") & "</TD></TR><TR><TD COLSPAN=2><B>---------------------------------------------------------------------------------------</B></TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>How they heard about Artisan Renovations: </B></TD><TD WIDTH=450>" &  Request.Form("Source") & "</TD></TR>"
sHTML = sHTML & "<TR><TD WIDTH=150 VALIGN=TOP><B>Customer was referred by: </B></TD><TD WIDTH=450>" &  Request.Form("referred") & "</TD></TR>"
sHTML = sHTML & "</TABLE></BODY></HTML>"

'Dim recaptcha_secret, sendstring, objNewMail 
' Secret key
        'recaptcha_secret = "6Lf5Ax0UAAAAAOU7KyC9Mqo5Ab_aFUnwSDIW-cwS"
        'sendstring = "https://www.google.com/recaptcha/api/siteverify?secret=" & recaptcha_secret & "&response=" & Request.form("g-recaptcha-response")
Dim objNewMail
Set objNewMail = Server.CreateObject("CDO.Message") 
Set objConfig = Server.CreateObject("CDO.Configuration") 

'Configuration: 
objConfig.Fields(cdoSendUsingMethod) = cdoSendUsingPort
objConfig.Fields(cdoSMTPServer)="smtp.1and1.com" 
objConfig.Fields(cdoSMTPServerPort)=25 
objConfig.Fields(cdoSMTPAuthenticate)=cdoBasic 
objConfig.Fields(cdoSendUserName) = "jschertz@ar-remodeling.com"
objConfig.Fields(cdoSendPassword) = "webdiva"
'Update configuration 
objConfig.Fields.Update 
Set objNewMail.Configuration = objConfig 

objNewMail.From = Request.Form("email")
'objNewMail.From = "jschertz1@comcast.net"
objNewMail.To = "jschertz1@comcast.net, SSchertz@ar-remodeling.com"
objNewMail.Subject = "Artisan Renovations Bid Request"
objNewMail.HTMLBody = sHTML
'objNewMail.BodyFormat = Mail.MailFormat.Html
objNewMail.Send
'objNewMail.MailFormat = 0
'objNewMail.BodyFormat = 0

Set objNewMail = Nothing
Set objConfig=Nothing 
%>
<!-- ADDED DOCTYPE -->
<!DOCTYPE html>
<HTML>

<head>
<title>Artisan Renovations - Request a Bid</title>

 <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" type="text/css" href="css/bootstrap.css">
    <link rel="stylesheet" type="text/css" href="css/font-awesome.css">
    <link rel='stylesheet' id='camera-css' href='css/camera.css' type='text/css' media='all'>

    <link rel="stylesheet" type="text/css" href="css/slicknav.css">
    <link rel="stylesheet" href="css/prettyPhoto.css" type="text/css" media="screen" title="prettyPhoto main stylesheet" />
    <link rel="stylesheet" type="text/css" href="css/style.css">
  
    <script type="text/javascript" src="js/jquery-1.8.3.min.js"></script>

    <link href='http://fonts.googleapis.com/css?family=Roboto:400,300,700|Open+Sans:700' rel='stylesheet' type='text/css'>
    <script type="text/javascript" src="js/jquery.mobile.customized.min.js"></script>
    <script type="text/javascript" src="js/jquery.easing.1.3.js"></script>
    <script type="text/javascript" src="js/camera.min.js"></script>
    <script type="text/javascript" src="js/myscript.js"></script>
    <script src="js/sorting.js" type="text/javascript"></script>
    <script src="js/jquery.isotope.js" type="text/javascript"></script>
    
  <script src="https://www.google.com/recaptcha/api.js" async defer></script>

</HEAD>
<body>


        <div class="headerLine">
            <div id="menuF" class="default">
                <div class="container">
                    <div class="row">
                        <div class="logo col-md-4">
                            <div>
                                <a href="#">
                                    <!--
                                    <img src="images/logo.png"> -->
                                    <img src="images/AR-text.gif">
                                </a>
                            </div>
                        </div>
                        <div class="col-md-8">
                            <div class="navmenu" style="text-align: center;">
                                <ul id="menu">
                                    <li class="active"><a href="index.html#home">Home</a></li>
                                    <li><a href="index.html#about">About</a></li>
                                    <li><a href="index.html#project">Projects</a></li>
                                    <li><a href="index.html#news">Design/Build</a></li>
                                    <li class="last"><a href="index.html#contact">Bid Request</a></li>

                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="container">
                <div class="row wrap">
                    <div class="col-md-12 gallery2">&nbsp;
                  </div>
                </div>
            </div>
        </div>
  <div class="container">
            <div class="row">
                <div class="col-md-12 Confirmation">
                    <h4>Thank you for your interest in Artisan Renovations.</h4>
                    <h4>Your request for a project bid has been submitted.</h4>
                    <h4>We will contact you soon to schedule a date and time to meet with you.</h4>
                    <br />
<h4><a href="index.html#home">Home</a></h4>
                </div>
            </div>
        </div>


<br/>

    <script src="js/jquery.prettyPhoto.js" type="text/javascript" charset="utf-8"></script>
    <script src="js/bootstrap.min.js"></script>
    <script src="js/jquery.slicknav.js"></script>
    <script>
        $(document).ready(function () {
            $(".bhide").click(function () {
                $(".hideObj").slideDown();
                $(this).hide(); //.attr()
                return false;
            });
            $(".bhide2").click(function () {
                $(".container.hideObj2").slideDown();
                $(this).hide(); // .attr()
                return false;
            });

            $('.heart').mouseover(function () {
                $(this).find('i').removeClass('fa-heart-o').addClass('fa-heart');
            }).mouseout(function () {
                $(this).find('i').removeClass('fa-heart').addClass('fa-heart-o');
            });

            function sdf_FTS(_number, _decimal, _separator) {
                var decimal = (typeof (_decimal) != 'undefined') ? _decimal : 2;
                var separator = (typeof (_separator) != 'undefined') ? _separator : '';
                var r = parseFloat(_number)
                var exp10 = Math.pow(10, decimal);
                r = Math.round(r * exp10) / exp10;
                rr = Number(r).toFixed(decimal).toString().split('.');
                b = rr[0].replace(/(\d{1,3}(?=(\d{3})+(?:\.\d|\b)))/g, "\$1" + separator);
                r = (rr[1] ? b + '.' + rr[1] : b);

                return r;
            }

            setTimeout(function () {
                $('#counter').text('0');
                $('#counter1').text('0');
                $('#counter2').text('0');
                setInterval(function () {

                    var curval = parseInt($('#counter').text());
                    var curval1 = parseInt($('#counter1').text().replace(' ', ''));
                    var curval2 = parseInt($('#counter2').text());
                    if (curval <= 707) {
                        $('#counter').text(curval + 1);
                    }
                    if (curval1 <= 12280) {
                        $('#counter1').text(sdf_FTS((curval1 + 20), 0, ' '));
                    }
                    if (curval2 <= 245) {
                        $('#counter2').text(curval2 + 1);
                    }
                }, 2);

            }, 500);
        });
    </script>
    <script type="text/javascript">
        jQuery(document).ready(function () {
            jQuery('#menu').slicknav();

        });
    </script>

    <script type="text/javascript">
        $(document).ready(function () {

            var $menu = $("#menuF");

            $(window).scroll(function () {
                if ($(this).scrollTop() > 100 && $menu.hasClass("default")) {
                    $menu.fadeOut('fast', function () {
                        $(this).removeClass("default")
                               .addClass("fixed transbg")
                               .fadeIn('fast');
                    });
                } else if ($(this).scrollTop() <= 100 && $menu.hasClass("fixed")) {
                    $menu.fadeOut('fast', function () {
                        $(this).removeClass("fixed transbg")
                               .addClass("default")
                               .fadeIn('fast');
                    });
                }
            });
        });
        //jQuery
    </script>
    <script>
        /*menu*/
        function calculateScroll() {
            var contentTop = [];
            var contentBottom = [];
            var winTop = $(window).scrollTop();
            var rangeTop = 200;
            var rangeBottom = 500;
            $('.navmenu').find('a').each(function () {
                contentTop.push($($(this).attr('href')).offset().top);
                contentBottom.push($($(this).attr('href')).offset().top + $($(this).attr('href')).height());
            })
            $.each(contentTop, function (i) {
                if (winTop > contentTop[i] - rangeTop && winTop < contentBottom[i] - rangeBottom) {
                    $('.navmenu li')
                    .removeClass('active')
                    .eq(i).addClass('active');
                }
            })
        };

        $(document).ready(function () {
            calculateScroll();
            $(window).scroll(function (event) {
                calculateScroll();
            });
            $('.navmenu ul li a').click(function () {
                $('html, body').animate({ scrollTop: $(this.hash).offset().top - 80 }, 800);
                return false;
            });
        });
    </script>

</body>
</html>