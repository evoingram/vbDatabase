<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="production.aspx.vb" Inherits="WebApp1._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <script>
        function hideMainFrameContent() {
            document.getElementByClass("mainFrameContent").style.visibility = "hidden";
        }
        function firstLoad() {
            document.getElementById("main").style.visibility = "visible";
        }
        function toggleMainFrameVis(pageID) {
            hideMainFrameContent();
            if () {
                document.getElementById(pageID).style.visibility = "hidden";
            } else (){
                document.getElementById(pageID).style.visibility = "visible";

            }
        }
        window.onload = firstload();
    </script>
    <div>
        <div class="row">

            <div class="col-md-3">
                <h1>Welcome to AQC Transcript Production.</h1>
                <p><a href="#" class="btn btn-primary btn-lg">Check FTP
                    <br />
                    for New Files</a></p>
                <p><a href="#newJob" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis('newJob');">Enter New Job</a></p>
                <p><a href="#priceQuote" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis('priceQuote');">Price Quote</a></p>
                <p><a href="#schedule" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Scheduling</a></p>
                <p><a href="#production" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis('production');">Transcript Production</a></p>
                <p><a href="#untagged" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Untagged Comms</a></p>
                <p><a href="#finances" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Finances</a></p>
                <p><a href="#searchJobs" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Search Jobs</a></p>
                <p><a href="#searchCitations" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Search Citations</a></p>
                <p><a href="#romanNumeral" class="btn btn-primary btn-lg" onclick="toggleMainFrameVis();">Roman Numeral
                    <br />
                    Converter &raquo;</a></p>
            </div>

            <div id="mainFrame" class="col-md-9">

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




                <div id="main" class="mainFrameContent">

                <div class="jumbotron">
                    <h1>Welcome to AQC Transcript Production.</h1>
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



            </div>


        </div>

    </div>
</asp:Content>
