<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="production.aspx.vb" Inherits="WebApp1._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">


    <div>
        <div class="row">

            <div id="sidebar" class="col-md-3">
                <h1>AQC Transcript Production</h1>
                <p>
                    <a href="#" class="btn btn-primary btn-lg">Check FTP
                    <br />
                        for New Files</a>
                </p>
                <p><a href="#newJob" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis('newJob', 'mainFrameContent')">Enter New Job</a></p>
                <p><a href="#priceQuote" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis('priceQuote', 'mainFrameContent')">Price Quote</a></p>
                <p><a href="#schedule" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Job Status</a></p>
                <p><a href="#production" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis('production', 'mainFrameContent')">Transcript Production</a></p>
                <p><a href="#untagged" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Untagged Comms</a></p>
                <p><a href="#finances" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Finances</a></p>
                <p><a href="#searchJobs" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Search Jobs</a></p>
                <p><a href="#searchCitations" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Search Citations</a></p>
                <p>
                    <a href="#romanNumeral" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Roman Numeral
                    <br />
                        Converter &raquo;</a>
                </p>
                    <p><a href="#userAdmin" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">User Admin</a></p>



            </div>

            <div id="mainFrame" class="col-md-9">

                <div id="main" class="mainFrameContent" onload="loadMain()">

                    <div class="row">

                        <div class="jumbotron">
                            <h1>Welcome to AQC Transcript Production.</h1>
                            <p class="lead">ASP.NET is a free web framework for building great Web sites and Web applications using HTML, CSS and JavaScript.</p>
                            <p><a href="http://www.asp.net" class="btn btn-primary btn-lg">Learn more &raquo;</a></p>
                        </div>

                    </div>

                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                            <p>
                                ASP.NET Web Forms lets you build dynamic websites using a familiar drag-and-drop, event-driven model.
                A design surface and hundreds of controls and components let you rapidly build sophisticated, powerful UI-driven sites with data access.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301948">Learn more &raquo;</a>
                            </p>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                            <p>
                                NuGet is a free Visual Studio extension that makes it easy to add, remove, and update libraries and tools in Visual Studio projects.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301949">Learn more &raquo;</a>
                            </p>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>
                            <p>
                                You can easily find a web hosting company that offers the right mix of features and price for your applications.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301950">Learn more &raquo;</a>
                            </p>
                        </div>
                    </div>


                </div>

                <div id="priceQuote" class="mainFrameContent">

                    <div class="jumbotron">
                        <h1>Get your price quote here.</h1>
                        <p class="lead">ASP.NET is a free web framework for building great Web sites and Web applications using HTML, CSS and JavaScript.</p>
                        <p><a href="http://www.asp.net" class="btn btn-primary btn-lg">Learn more &raquo;</a></p>
                    </div>

                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                            <p>
                                ASP.NET Web Forms lets you build dynamic websites using a familiar drag-and-drop, event-driven model.
                A design surface and hundreds of controls and components let you rapidly build sophisticated, powerful UI-driven sites with data access.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301948">Learn more &raquo;</a>
                            </p>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                            <p>
                                NuGet is a free Visual Studio extension that makes it easy to add, remove, and update libraries and tools in Visual Studio projects.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301949">Learn more &raquo;</a>
                            </p>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>
                            <p>
                                You can easily find a web hosting company that offers the right mix of features and price for your applications.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301950">Learn more &raquo;</a>
                            </p>
                        </div>
                    </div>

                </div>


                <div id="production" class="mainFrameContent">

                    <div class="jumbotron">
                        <h1>Production.</h1>
                        <p class="lead">ASP.NET is a free web framework for building great Web sites and Web applications using HTML, CSS and JavaScript.</p>
                        <p><a href="http://www.asp.net" class="btn btn-primary btn-lg">Learn more &raquo;</a></p>
                    </div>

                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                            <p>
                                ASP.NET Web Forms lets you build dynamic websites using a familiar drag-and-drop, event-driven model.
                A design surface and hundreds of controls and components let you rapidly build sophisticated, powerful UI-driven sites with data access.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301948">Learn more &raquo;</a>
                            </p>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                            <p>
                                NuGet is a free Visual Studio extension that makes it easy to add, remove, and update libraries and tools in Visual Studio projects.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301949">Learn more &raquo;</a>
                            </p>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>
                            <p>
                                You can easily find a web hosting company that offers the right mix of features and price for your applications.
                            </p>
                            <p>
                                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301950">Learn more &raquo;</a>
                            </p>
                        </div>
                    </div>

                </div>



                <div id="newJob" class="mainFrameContent">

                    <div class="jumbotron">
                        <h1>Production.</h1>
                    </div>

                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>

                        </div>
                    </div>

                </div>


                <div id="schedule" class="mainFrameContent">

                    <div class="jumbotron">
                        <h1>Production.</h1>
                    </div>

                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>

                        </div>
                    </div>

                </div>

                <div id="untagged" class="mainFrameContent">

                    <div class="jumbotron">
                        <h1>Production.</h1>
                    </div>

                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>

                        </div>
                    </div>

                </div>


                <div id="finances" class="mainFrameContent">

                    <div class="jumbotron">
                        <h1>Production.</h1>
                    </div>

                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>

                        </div>
                    </div>

                </div>

                <div id="searchJobs" class="mainFrameContent">
                    <div class="jumbotron">
                        <h1>Production.</h1>
                    </div>
                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>
                        </div>
                    </div>

                </div>

                <div id="searchCitations" class="mainFrameContent">
                    <div class="jumbotron">
                        <h1>Production.</h1>
                    </div>
                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>
                        </div>
                    </div>
                </div>
                <div id="romanNumeral" class="mainFrameContent">
                    <div class="jumbotron">
                        <h1>Production.</h1>
                    </div>
                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>
                        </div>
                    </div>
                </div>

                <div id="userAdmin" class="mainFrameContent">
                    <div class="jumbotron">
                        <h1>Production.</h1>
                    </div>
                    <div class="row">
                        <div class="col-md-4">
                            <h2>Transcript Production</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Get more libraries</h2>
                        </div>
                        <div class="col-md-4">
                            <h2>Web Hosting</h2>
                        </div>
                    </div>
                </div>

            </div>


        </div>


    </div>


    <script>
        function toggleMainFrameVis(pageID, className) {
            let d;
            let x;
            // if baseClass of any of the three sections is showing, set to hide
            // loops through all of the divs labeled with the class name.
            if (document.getElementsByClassName(className).display == "block") {
                document.getElementById(pageID).style.display = "none";
                d = document.getElementsByClassName(className);
                for (x = 0; x < d.length; x++) {
                    d[x].style.display = 'none';
                }
            }
            // if page w/ page ID is showing, hide it
            if (document.getElementById(pageID).style.display == "block") {
                d = document.getElementById(pageID);
                d.style.display = "none";
            }
            else { // if not showing, hide every one on page, and then show the correct one
                d = document.getElementsByClassName(className);
                for (x = 0; x < d.length; x++) {
                    d[x].style.display = 'none';
                }
                d = document.getElementById(pageID);
                d.style.display = "block";
            }

        }
        // when dom is ready, it loads main section as default
        $(document).ready(function () {
            let d = document.getElementById("main");
            d.style.display = "block";
        });
    </script>
</asp:Content>
