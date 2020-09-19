from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pyodbc
import time
import os



# Define Today's Date and Time
now = datetime.today().strftime('%d/%m/%Y')

# Connect To SQL
conn = pyodbc.connect('Driver={SQL Server};' 'Server=SKW-TBSQL;''Database=BIA_REPORTS_2019v2;' 'Trusted_Connection=yes;')
cursor = conn.cursor()


rows = cursor.execute("""SELECT dbo.PropertyDatabase.PropertyName, dbo.PropertyDatabase.PropertyType, dbo.PropertyDatabase.PropertyID, dbo.PropertyDatabase.ResortID, dbo.PropertyDatabase.Resort, 
                         dbo.PropertyDatabase.TransferLocationID, dbo.PropertyDatabase.TransferLocation, dbo.PropertyDatabase.CountryID, dbo.PropertyDatabase.Country, dbo.PropertyDatabase.PropertyRating, 
                         dbo.PropertyDatabase.Supplier, dbo.PropertyDatabase.checkInTime, dbo.PropertyDatabase.checkOutTime, dbo.PropertyDatabase.DMSItinerariesConfirmedAccommodationText, 
                         dbo.PropertyDatabase.DMSItinerariesOptionAccommodationText, dbo.PropertyDatabase.DMSQuoteColumn1, dbo.PropertyDatabase.DMSQuoteColumn2, dbo.PropertyDatabase.OfficalRating, 
                         dbo.PropertyDatabase.SubResort, dbo.PropertyDatabase.SkiworldWebsite, dbo.PropertyDatabase.PropertyPhotos, dbo.PropertyDatabase.SalesQuote, dbo.PropertyDatabase.WebsiteDescription, 
                         dbo.PropertyDatabase.OutOfHoursCheckIn, dbo.PropertyDatabase.SecurityDeposit, dbo.PropertyDatabase.TouristTax, dbo.PropertyDatabase.WebsiteLocationDescription, dbo.PropertyDatabase.GoogleMaps, 
                         dbo.PropertyDatabase.DistancetoSkiHireShop, dbo.PropertyDatabase.DistanceToBusStop, dbo.PropertyDatabase.DistanceToSkiSchool, dbo.PropertyDatabase.DistanceToSupermarket, 
                         dbo.PropertyDatabase.DistanceToTownCentre, dbo.PropertyDatabase.DrivingDirections, dbo.PropertyDatabase.OnSiteParking, dbo.PropertyDatabase.ReceptionContact, dbo.PropertyDatabase.ReceptionTimes, 
                         dbo.PropertyDatabase.LiftPassPickupInformation, dbo.PropertyDatabase.SkiHirePickupInformation, dbo.PropertyDatabase.SkiHireShopName, dbo.PropertyDatabase.SkiHireGoogleMaps, 
                         dbo.PropertyDatabase.RepService, dbo.PropertyDatabase.SalesOfSkiPacks, dbo.PropertyDatabase.StandardTransfers, dbo.PropertyDatabase.IncludedServices, dbo.PropertyDatabase.StandardBoardBasis, 
                         dbo.PropertyDatabase.StandardMealDescription, dbo.PropertyDatabase.ExtraBoardSupplements, dbo.PropertyDatabase.RoomFacilities, dbo.PropertyDatabase.WIFI, dbo.PropertyDatabase.PayableServices, 
                         dbo.PropertyDatabase.Bars, dbo.PropertyDatabase.Restaurants, dbo.PropertyDatabase.SwimmingPool, dbo.PropertyDatabase.HotTubOrWhirlpool, dbo.PropertyDatabase.SaunaOrSteamRoom, 
                         dbo.PropertyDatabase.Gym, dbo.PropertyDatabase.Spa, dbo.PropertyDatabase.KitStorageFacilities, dbo.PropertyDatabase.MinibusOrShuttleService, dbo.PropertyDatabase.Pets, 
                         dbo.PropertyDatabase.ChildandInfantsServices, dbo.PropertyDatabase.ChildcareFacilities, dbo.PropertyDatabase.SupplierWebsite, dbo.PropertyDatabase.SupplierContact, 
                         dbo.PropertyDatabase.SupplierContactTitle, dbo.PropertyDatabase.SupplierContactEmail, dbo.PropertyDatabase.SupplierContactPhone, dbo.PropertyDatabase.SupplierBookingMethod, 
                         dbo.PropertyDatabase.SupplierReservationPhone, dbo.PropertyDatabase.SupplierReservationsEmail, dbo.PropertyDatabase.SupplierBankDetails, dbo.PropertyDatabase.SupplierPaymentTerms, 
                         dbo.PropertyDatabase.SupplierCancellationPolicy, dbo.PropertyDatabase.DistanceToPiste, dbo.PropertyDatabase.DistanceToLift, dbo.ResortDatabase.ResortName, dbo.ResortDatabase.ResortID AS Expr1, 
                         dbo.ResortDatabase.Country AS Expr2, dbo.ResortDatabase.CountryID AS Expr3, dbo.ResortDatabase.SkiworldResortWebpage, dbo.ResortDatabase.WebsiteResortID, dbo.ResortDatabase.Resortheight, 
                         dbo.ResortDatabase.Highestlift, dbo.ResortDatabase.Transfertimes, dbo.ResortDatabase.Guideliftpassprice, dbo.ResortDatabase.Totalpiste, dbo.ResortDatabase.Noofslopes, dbo.ResortDatabase.Nooflifts, 
                         dbo.ResortDatabase.Eurostar, dbo.ResortDatabase.Overview, dbo.ResortDatabase.Skiing, dbo.ResortDatabase.Activities, dbo.ResortDatabase.Apres, dbo.ResortDatabase.Dining, dbo.ResortDatabase.Family, 
                         dbo.ResortDatabase.SkiHireWhatsincluded, dbo.ResortDatabase.SkiHireBronzePackage, dbo.ResortDatabase.SkiHireSilverPackage, dbo.ResortDatabase.SkiHireGoldPackage, 
                         dbo.ResortDatabase.SkiHirePlatinumPackage, dbo.ResortDatabase.LiftPassWhatsincluded, dbo.ResortDatabase.LiftPassInsiderTips, dbo.ResortDatabase.Fire, dbo.ResortDatabase.Police, 
                         dbo.ResortDatabase.Ambulance, dbo.ResortDatabase.TouristOffice, dbo.ResortDatabase.MountainRescue, dbo.ResortDatabase.MedicalCentreLocationandDirections, 
                         dbo.ResortDatabase.SkiworldEmergencyContactInformation, dbo.ResortDatabase.SalesQuote AS Expr4, dbo.ResortDatabase.SkiingAdvanced, dbo.ResortDatabase.Snowboarders, 
                         dbo.ResortDatabase.BestForFamily, dbo.ResortDatabase.BestForGroup, dbo.ResortDatabase.BestForBeginners, dbo.ResortDatabase.BestForIntermediates, dbo.ResortDatabase.BestForAdvanced, 
                         dbo.ResortDatabase.BestForSnowboarders, dbo.ResortDatabase.KeyFeatureShortTransfer, dbo.ResortDatabase.KeyFeatureGoodForApres, dbo.ResortDatabase.KeyFeatureFineDining, 
                         dbo.ResortDatabase.KeyFeatureActivitiesForNonSkiers, dbo.ResortDatabase.KeyFeatureTraditionalVillage, dbo.ResortDatabase.KeyFeatureGoodForSelfDrive, dbo.ResortDatabase.KeyFeatureEurostarTrain, 
                         dbo.ResortDatabase.KeyFeatureLuxuryAccommodation, dbo.ResortDatabase.KeyFeatureGoodForSkiInSkiOut, dbo.ResortDatabase.KeyFeatureValueForMoneyResort, dbo.ResortDatabase.KeyFeatureLuxury, 
                         dbo.ResortDatabase.SkiAreaHighResortSnowSure, dbo.ResortDatabase.SkiAreaLargeSkiArea, dbo.ResortDatabase.SkiAreaGlacier, dbo.ResortDatabase.SkiAreaGoodForOffPiste, 
                         dbo.ResortDatabase.SkiingBeginners, dbo.ResortDatabase.SkiingIntermediates
                        FROM dbo.PropertyDatabase 
                        LEFT OUTER JOIN
                        dbo.ResortDatabase ON dbo.PropertyDatabase.ResortID = dbo.ResortDatabase.ResortID
                        WHERE (dbo.PropertyDatabase.Supplier <> 'Committed Costs GBP') AND (dbo.PropertyDatabase.Supplier <> 'EUR Supplier')""")

for row in rows:
    if str(row[4]) == 'Tignes':
        PropertyName = str(row[0])
        PropertyType = str(row[1])
        PropertyID = str(row[2])
        ResortID = str(row[3])
        Resort = str(row[4])
        TransferLocationID = str(row[5])
        TransferLocation = str(row[6])
        CountryID = str(row[7])
        Country = str(row[8])
        PropertyRating = str(row[9])
        Supplier = str(row[10])
        checkInTime = str(row[11])
        checkOutTime = str(row[12])
        DMSItinerariesConfirmedAccommodationText = str(row[13])
        DMSItinerariesOptionAccommodationText = str(row[14])
        DMSQuoteColumn1 = str(row[15])
        DMSQuoteColumn2 = str(row[16])
        OfficalRating = str(row[17])
        SubResort = str(row[18])
        SkiworldWebsite = str(row[19])
        PropertyPhotos = str(row[20])
        SalesQuote = str(row[21])
        WebsiteDescription = str(row[22])
        OutOfHoursCheckIn = str(row[23])
        SecurityDeposit = str(row[24])
        TouristTax = str(row[25])
        WebsiteLocationDescription = str(row[26])
        GoogleMaps = str(row[27])
        DistancetoSkiHireShop = str(row[28])
        DistanceToBusStop = str(row[29])
        DistanceToSkiSchool = str(row[30])
        DistanceToSupermarket = str(row[31])
        DistanceToTownCentre = str(row[32])
        DrivingDirections = str(row[33])
        OnSiteParking = str(row[34])
        ReceptionContact = str(row[35])
        ReceptionTimes = str(row[36])
        LiftPassPickupInformation = str(row[37])
        SkiHirePickupInformation = str(row[38])
        SkiHireShopName = str(row[39])
        SkiHireGoogleMaps = str(row[40])
        RepService = str(row[41])
        SalesOfSkiPacks = str(row[42])
        StandardTransfers = str(row[43])
        IncludedServices = str(row[44])
        StandardBoardBasis = str(row[45])
        StandardMealDescription = str(row[46])
        ExtraBoardSupplements = str(row[47])
        RoomFacilities = str(row[48])
        WIFI = str(row[49])
        PayableServices = str(row[50])
        Bars = str(row[51])
        Restaurants = str(row[52])
        SwimmingPool = str(row[53])
        HotTubOrWhirlpool = str(row[54])
        SaunaOrSteamRoom = str(row[55])
        Gym = str(row[56])
        Spa = str(row[57])
        KitStorageFacilities = str(row[58])
        MinibusOrShuttleService = str(row[59])
        Pets = str(row[60])
        ChildandInfantsServices = str(row[61])
        ChildcareFacilities = str(row[62])
        SupplierWebsite = str(row[63])
        SupplierContact = str(row[64])
        SupplierContactTitle = str(row[65])
        SupplierContactEmail = str(row[66])
        SupplierContactPhone = str(row[67])
        SupplierBookingMethod = str(row[68])
        SupplierReservationPhone = str(row[69])
        SupplierReservationsEmail = str(row[70])
        SupplierBankDetails = str(row[71])
        SupplierPaymentTerms = str(row[72])
        SupplierCancellationPolicy = str(row[73])
        DistanceToPiste = str(row[74])
        DistanceToLift = str(row[75])
        ResortName = str(row[76])
        ResortID2 = str(row[77])
        Country2 = str(row[78])
        CountryID2 = str(row[79])
        SkiworldResortWebpage = str(row[80])
        WebsiteResortID = str(row[81])
        Resortheight = str(row[82])
        Highestlift = str(row[83])
        Transfertimes = str(row[84])
        Guideliftpassprice = str(row[85])
        Totalpiste = str(row[86])
        Noofslopes = str(row[87])
        Nooflifts = str(row[88])
        Eurostar = str(row[89])
        Overview = str(row[90])
        Skiing = str(row[91])
        Activities = str(row[92])
        Apres = str(row[93])
        Dining = str(row[94])
        Family = str(row[95])
        SkiHireWhatsincluded = str(row[96])
        SkiHireBronzePackage = str(row[97])
        SkiHireSilverPackage = str(row[98])
        SkiHireGoldPackage = str(row[99])
        SkiHirePlatinumPackage = str(row[100])
        LiftPassWhatsincluded = str(row[101])
        LiftPassInsiderTips = str(row[102])
        Fire = str(row[103])
        Police = str(row[104])
        Ambulance = str(row[105])
        TouristOffice = str(row[106])
        MountainRescue = str(row[107])
        MedicalCentreLocationandDirections = str(row[108])
        SkiworldEmergencyContactInformation = str(row[109])
        ResortSalesQuote = str(row[110])
        SkiingAdvanced = str(row[111])
        Snowboarders = str(row[112])
        BestForFamily = str(row[113])
        BestForGroup = str(row[114])
        BestForBeginners = str(row[115])
        BestForIntermediates = str(row[116])
        BestForAdvanced = str(row[117])
        BestForSnowboarders = str(row[118])
        KeyFeatureShortTransfer = str(row[119])
        KeyFeatureGoodForApres = str(row[120])
        KeyFeatureFineDining = str(row[121])
        KeyFeatureActivitiesForNonSkiers = str(row[122])
        KeyFeatureTraditionalVillage = str(row[123])
        KeyFeatureGoodForSelfDrive = str(row[124])
        KeyFeatureEurostarTrain = str(row[125])
        KeyFeatureLuxuryAccommodation = str(row[126])
        KeyFeatureGoodForSkiInSkiOut = str(row[127])
        KeyFeatureValueForMoneyResort = str(row[128])
        KeyFeatureLuxury = str(row[129])
        SkiAreaHighResortSnowSure = str(row[130])
        SkiAreaLargeSkiArea = str(row[131])
        SkiAreaGlacier = str(row[132])
        SkiAreaGoodForOffPiste = str(row[133])
        SkiingBeginners = str(row[134])
        SkiingIntermediates = str(row[135])
        print(PropertyName)

        Filename = """C:/Users/products/Desktop/Python Scripts/Essential Arrival Docs/""" + ResortName + """ - """ + PropertyName + """.html"""
        f = open(Filename, 'w')

        HTML = """<html>
        <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        </head>
        <body>

        <style type="text/css">/* Email client specific styles */
               
                img {
                    border: 0;
                    height: auto;
                    line-height: 100%;
                    outline: none;
                    text-decoration: none;
                }
               
        @font-face {
            font-family: 'SourceSansPro';
            src: url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-Regular.eot');
            src: url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-Regular.eot?#iefix') format('embedded-opentype'),
                url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-Regular.woff') format('woff'),
                url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-Regular.ttf') format('truetype');
            font-weight: normal;
            font-style: normal;
        }

        @font-face {
            font-family: 'SourceSansPro';
            src: url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-Bold.eot');
            src: url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-Bold.eot?#iefix') format('embedded-opentype'),
                url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-Bold.woff') format('woff'),
                url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-Bold.ttf') format('truetype');
            font-weight: bold;
            font-style: normal;
        }

        @font-face {
            font-family: 'SourceSansPro';
            src: url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-It.eot');
            src: url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-It.eot?#iefix') format('embedded-opentype'),
                url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-It.woff') format('woff'),
                url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-It.ttf') format('truetype');
            font-weight: normal;
            font-style: italic;
        }

        @font-face {
            font-family: 'SourceSansPro';
            src: url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-BoldIt.eot');
            src: url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-BoldIt.eot?#iefix') format('embedded-opentype'),
                url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-BoldIt.woff') format('woff'),
                url('http://marketingradarltd.github.io/SkiWorld/fonts/SourceSansPro-BoldIt.ttf') format('truetype');
            font-weight: bold;
            font-style: italic;
        }

        body {
        font-family: Arial, Helvetica, sans-serif;
        font-size: 15px;
        line-height: 1.5em;
        padding: 0px 10px 0px 10px;
        }   

        h1 {
            font-family: Arial, Helvetica, sans-serif; 
            color: #042859;
            font-size: 24px;
            line-height: 1.5em;
            padding: 0px;
            margin-top: 30px;
            text-align:left;
        }

        h2 {
            font-family: Arial, Helvetica, sans-serif; 
            color: #042859;
            font-size: 20px;
            line-height: 1.5em;
            padding: 0px;
            margin: 0px;
        }
        h3 {
            font-family: Arial, Helvetica, sans-serif; 
            color: #042859;
            font-size: 18px;
            line-height: 1.5em;
            padding: 0px;
            margin: 0px;
        }
        h4 {
            font-family: Arial, Helvetica, sans-serif; 
            color: #042859;
            font-size: 16px;
            line-height: 1.5em;
            padding: 0px 10px 0px 0px;
        }

            
        hr {
            height: 3px;
            background-color:#a6a6a6;
            margin: 30px 0px;
        } 
        table, table td {
            border-collapse: collapse;
            font-family: Arial, Helvetica, sans-serif;
            font-size: 15px;
            line-height: 1.5em;
        }
        td {
            vertical-align:top;
            padding: 5px 0px;
        }
        a, a:hover, a:visited {
        color:#349bb3;
        font-weight:bold;
        text-decoration:underline;
        }
        li {
          margin-bottom:15px;
        }


                
        </style>

        <div style="width:600px;margin-left:auto;margin-right:auto">
        <div><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=Essential-Arrival-Info-Header-Image.jpg" width="600px"/></div>

        <h4>We'd like your holiday to run as smoothly as waxed skis (or boards) on fresh snow.</h4>

        <p>We have prepared this information to make sure you have everything you need for your ski holiday with us.
        Please read through and don't hesitate to call if you need anything. Don't forget you can also view your
        documents and enter passport details by logging on to <a href="https://www.skiworld.co.uk/my-skiworld/login">My Skiworld</a>. If you have not already, please check your
        documents to make sure that we have your names as they appear on your passports, and that you have given us
        your emergency contact number in case we need to contact you whilst you are in resort regarding your
        transport arrangements.
        </p>
        <p>
        This content is split into two sections. First you'll find essential arrival information, and then further below you'll find our <a href="#resort-guide">resort guide</a>, which gives some insight into the resort, non-ski activities available, places to eat and some apres spots. We hope you find it helpful.
        </p>

        <hr/>


        <table>
        <tr>
        <td style="width:80px"></td>
        <td><h1>Essential Arrival Information</h1></td>
        </tr>

        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-search.jpg" width="60px"/></td>
        <td><h2>Contents</h2></td>
        </tr>
        <tr>
        <td>
        <a href="#property-info">Property Information</a><br/>
        <a href="#arrival-info">Arrival and Contact Information</a><br/>
        <a href="#ski-essentials">Ski Essentials</a><br/>
        <a href="#self-drive-info">Self-Drive and Location Information</a><br/>
        <a href="#medical-info">Resort Emergency/Medical</a><br/>
        <a href="#emergency-contact">Emergency Contact Information</a>
        </td>
        </tr>
        </table>

        <hr/>

        <a name="property-info"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-accom.jpg" width="60px"/></td>
        <td><h2>Property Information</h2></td>
        </tr>
        <tr>
        <td>
        <p>
        """ + PropertyName + """<br/>
        """ + Resort + """<br/>
        """ + Country + """<br/><br/>
        Please visit the Skiworld website for information about your property and the facilities here -<br/>
        <a href=""" + SkiworldWebsite + """>""" + PropertyName + """</a><br/><br/>
        For information about the skiing, activities, apres, dining and families please visit our resort page -<br/>
        <a href=""" + SkiworldResortWebpage + """>""" + Resort + """</a>
        </p>
        </td>
        </tr>
        </table>

        <hr/>

        <a name="arrival-info"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-phone.jpg" width="60px"/></td>
        <td><h2>Arrival and Contact Information</h2></td>
        </tr>
        <tr>
        <td>
        <table>
        <tr>
        <td style="width:200px">Check-In Time :</td>
        <td>""" + checkInTime + """</td>
        </tr>
        <tr>
        <td style="width:200px">Check-Out Time :</td>
        <td>""" + checkOutTime + """</td>
        </tr>
        <tr>
        <td style="width:200px">Reception Opening Time :</td>
        <td>""" + ReceptionTimes + """</td>
        </tr>
        <tr>
        <td style="width:200px">Reception Contact :</td>
        <td>""" + ReceptionContact + """</td>
        </tr>
        <tr>
        <td style="width:200px">Out Of Hours Check-In :</td>
        <td>""" + OutOfHoursCheckIn + """</td>
        </tr>
        </table>
        </td>
        </tr>
        <tr>
        <td></td>
        <td>&nbsp;</td>
        </tr>
        <tr>
        <td></td>
        <td>On arrival you will be asked to pay a security deposit and/or Taxe De Sejour. A tourist tax (Taxe de Sejour or Kurtaxe) is levied by local councils in European ski resorts to support the local tourism infrastructure and may include services such as ski buses, subsidised admission to amenities e.g. pools and ice rinks and events such as firework displays. The amount charged varies according to the standard and type of accommodation. It is also not paid by all skiers as it is age-dependent. Therefore it is not included in the basic price of your holiday as we do not have all skiers ages at the time of booking.<br/><br/>
        The amounts below are an approximation and can vary by property size and passenger age.</td>
        </tr>

        <tr>
        <td></td>
        <td>&nbsp;</td>
        </tr>
        <tr>
        <td>
        </td>
        <td>
        <table>
        <tr>
        <td style="width:200px">Security Deposit :</td>
        <td>""" + SecurityDeposit + """</td>
        </tr>
        <tr>
        <td style="width:200px">Tax De Sejour :</td>
        <td>""" + TouristTax + """</td>
        </tr>
        </table>
        </td>
        </tr>

        </table>

        <hr/>

        <a name="ski-essentials"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-ski.jpg" width="60px"/></td>
        <td><h2>Ski Essentials</h2></td>
        </tr>
        <tr>
        <td>If you have pre-booked ski essentials with Skiworld please see below for the relevant operational information.
        If you have not pre-booked your ski essentials please call 0330 102 8004. Skiworld will endeavour to price match on ski essentials and can offer a discount when booking your lift pass and ski hire together.<br/><br/>
        Please note ski essentials will no longer be available for purchase on the transfer to resort.</td>
        </tr>

        <tr>
        <td>
        </td>
        <td>
        <table>
        <tr>
        <td style="width:200px">Lift Pass Pickup :</td>
        <td>""" + LiftPassPickupInformation + """</td>
        </tr>
        <tr>
        <td style="width:200px">Ski Hire Pickup :</td>
        <td>""" + SkiHireShopName + SkiHirePickupInformation + """</td>
        </tr>
        <tr>
        <td style="width:200px">Lessons :</td>
        <td>All lesson information will be available on your itinerary, including start times and meeting points as these vary by provider and ski/board ability.</td>
        </tr>
        </table>
        </td>
        </tr>

        </table>

        <hr/>

        <a name="self-drive-info"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-self-drive.jpg" width="60px"/></td>
        <td><h2>Self-Drive and Location Information</h2></td>
        </tr>

        <tr>
        <td>
        <table>
        <tr>
        <td style="width:200px">Google Maps :</td>
        <td><a href=""" + GoogleMaps + """>""" + PropertyName + """</a></td>
        </tr>
        <tr>
        <td style="width:200px">Driving Directions :</td>
        <td>""" + DrivingDirections + """</td>
        </tr>

        <tr>
        <td style="width:200px">Parking :</td>
        <td>""" + OnSiteParking + """</td>
        </tr>

        <tr>
        <td style="width:200px">Distance To Skiing :</td>
        <td>""" + DistanceToPiste + """</td>
        </tr>

        <tr>
        <td style="width:200px">Distance To Town : :</td>
        <td>""" + DistanceToTownCentre + """</td>
        </tr>

        <tr>
        <td style="width:200px">Distance To Ski Hire :</td>
        <td>""" + DistancetoSkiHireShop + """</td>
        </tr>

        <tr>
        <td style="width:200px">Distance To Ski School :</td>
        <td>""" + DistanceToSkiSchool + """</td>
        </tr>

        <tr>
        <td style="width:200px">Distance To Supermarket :</td>
        <td>""" + DistanceToSupermarket + """</td>
        </tr>

        </table>
        </td>
        </tr>

        </table>

        <hr/>

        <a name="medical-info"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-exclamation.jpg" width="60px"/></td>
        <td><h2>Resort Emergency/Medical Information</h2></td>
        </tr>

        <tr>

        <td>
        <table>

        <tr>
        <td style="width:200px">Fire :</td>
        <td>""" + Fire + """</td>
        </tr>

        <tr>
        <td style="width:200px">Police :</td>
        <td>""" + Police + """</td>
        </tr>

        <tr>
        <td style="width:200px">Ambulance :</td>
        <td>""" + Ambulance + """</td>
        </tr>

        <tr>
        <td style="width:200px">Tourist Office :</td>
        <td>""" + TouristOffice + """</td>
        </tr>

        <tr>
        <td style="width:200px">Mountain Rescue :</td>
        <td>""" + MountainRescue + """</td>
        </tr>

        <tr>
        <td style="width:200px">Medical centres :</td>
        <td>""" + MedicalCentreLocationandDirections + """</td>
        </tr>

        </table>
        </td>
        </tr>

        </table>

        <hr/>

        <a name="emergency-contact"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-search.jpg" width="60px"/></td>
        <td><h2>Emergency Contact Information</h2></td>
        </tr>
        <tr>
        <td>

        """ + SkiworldEmergencyContactInformation + """</td>

        </tr>
        </table>

        <p>
        <strong>This information was correct to the best of our knowledge on """ + now + """.</strong> We do our very best to ensure that all information is as accurate as possible. However, we are sure that you will appreciate some things may change during the season without our knowledge - despite our efforts to check! Therefore, please consider the information above as our best guidelines as to distances, costs and what to expect. Please do let us know if you find anything that needs updating or would like us to provide more or different information.
        </p>

        <hr/>

        <br/>
        <br/>

        <a name="resort-guide"></a>
        <table>
        <tr>
        <td style="width:80px"></td>
        <td><h1>Resort Guide</h1></td>
        </tr>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-search.jpg" width="60px"/></td>
        <td><h2>Contents</h2></td>
        </tr>
        <tr>
        <td>
        <a href="#resort-overview">Overview</a><br/>
        <a href="#resort-skiing">Skiing</a><br/>
        <a href="#resort-activities">Activities</a><br/>
        <a href="#resort-apres">Apres </a><br/>
        <a href="#resort-dining">Dining</a><br/>
        <a href="#resort-families">Families</a><br/>
        <a href="#resort-pictures">Pictures</a>
        </td>
        </tr>
        </table>

        <hr/>

        <a name="resort-overview"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icons-piste.jpg" width="60px"/></td>
        <td><h2>Overview of Alpe d'Huez</h2></td>
        </tr>
        <tr>
        <td>
        """ + Overview + """
        <br/>
        <h3>Resort Facts and Stats</h3>
        <br/>
        <table>

        <tr>
        <td style="width:200px">Resort Height : </td>
        <td>""" + Resortheight + """</td>
        </tr>

        <tr>
        <td style="width:200px">Highest Lift : </td>
        <td>""" + Highestlift + """</td>
        </tr>

        <tr>
        <td style="width:200px">Transfer Time (approx) : </td>
        <td>""" + Transfertimes + """</td>
        </tr>

        <tr>
        <td style="width:200px">Total Piste : </td>
        <td>""" + Totalpiste + """</td>
        </tr>

        <tr>
        <td style="width:200px">Number of Slopes : </td>
        <td>""" + Noofslopes + """</td>
        </tr>

        <tr>
        <td style="width:200px">Number of Lifts : </td>
        <td>""" + Nooflifts + """</td>
        </tr>

        </table>
        </td>
        </tr>
        </table>




        <hr/>

        <a name="resort-skiing"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-ski.jpg" width="60px"/></td>
        <td><h2>Skiing and Boarding in Alpe d'Huez</h2></td>
        </tr>
        <tr>
        <td>
        """ + Skiing + """
        </td>
        </tr>
        </table>

        <hr/>

        <a name="resort-activities"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2">Need icon</td>
        <td><h2>Activities</h2></td>
        </tr>
        <tr>
        <td>
        """ + Activities + """
        </td>
        </tr>
        </table>

        <hr/>

        <a name="resort-apres"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2">Need icon</td>
        <td><h2>Apres</h2></td>
        </tr>
        <tr>
        <td>
        """ + Apres + """
        </td>
        </tr>
        </table>

        <hr/>

        <a name="resort-dining"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-dining.jpg" width="60px"/></td>
        <td><h2>Dining</h2></td>
        </tr>
        <tr>
        <td>
        """ + Dining + """
        </td>
        </tr>
        </table>

        <hr/>

        <a name="resort-Families"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-family.jpg" width="60px"/></td>
        <td><h2>Families</h2></td>
        </tr>
        <tr>
        <td>
        """ + Family + """
        </td>
        </tr>
        </table>

        <hr/>

        <a name="resort-pictures"></a>
        <table>
        <tr>
        <td style="width:80px" rowspan="2"><img src="https://skiworld.ontigerbay.co.uk/admin/CMS/resize/resize.ashx?f=icon-online.jpg" width="60px"/></td>
        <td><h2>Resort Pictures</h2></td>
        </tr>
        <tr>
        <td>
        <p>
        You can see pictures of the resort via our website:<br/><a href=""" + SkiworldResortWebpage + """>See """ + ResortName + """ on Skiworld.co.uk</a>.
        </p>
        </td>
        </tr>
        </table>

        <hr/>
        </div>

        </body></html>"""

        f.write(HTML)
        f.close()

    else:
        pass
        

