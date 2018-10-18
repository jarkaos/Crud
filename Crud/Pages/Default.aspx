<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <SharePoint:ScriptLink name="sp.js" runat="server" OnDemand="true" LoadAfterUI="true" Localizable="false" />
    <meta name="WebPartPageExpansion" content="full" />

    <!-- Add your CSS styles to the following file -->
    <link type="text/css" href="../Content/bootstrap.css" rel="stylesheet" />
    <link type="text/css" href="../Content/toastr.css" rel="stylesheet" />
    <link type="text/css" href="../Content/bootstrap-dialog.css" rel="stylesheet" />
    <link type="text/css" href="../Content/DataTables/css/select.bootstrap.min.css" rel="stylesheet" />
    <link href="../Content/DataTables/css/jquery.dataTables.min.css" rel="stylesheet" />
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/bootstrap.js"></script>
    <script type="text/javascript" src="../Scripts/bootstrap-dialog.js"></script>
    <script type="text/javascript" src="../Scripts/App.js"></script>
    <script type="text/javascript" src="../Scripts/toastr.min.js"></script>
    <script src="../Scripts/DataTables/jquery.dataTables.min.js"></script>
    <script type="text/javascript" src="../Scripts/DataTables/dataTables.select.min.js"></script>
    <script src="../Scripts/moment.min.js"></script>
</asp:Content>


<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">

   <div class="container">
        <h1><span class="label label-primary">Golf Bokningar</span></h1>
        <div id="toolbar">
            <button type="button" value="Bokningar" class="btn btn-info" onclick="Javascript: location.href = '../../Lists/GolfBokningar'"><span class='glyphicon glyphicon-upload'></span>Bokning List</button>
            <button type="button" class="btn btn-success" id="addNewBokning"><span class='glyphicon glyphicon-plus'></span>Ny bokning</button>
        </div>
        <p></p>
        <div id="BokningPanel">
            <table style="width: 100%;">
                <tr>
                    <td>
                        <div id="BokningsGrid" style="width: 100%"></div>
                    </td>
                </tr>
            </table>
        </div>
        <!-- Bootstrap Modal Dialog-->
        <div class="modal fade" id="myModalNorm" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                    <!-- Modal Header -->
                    <div class="modal-header">
                        <button type="button" class="close" data-dismiss="modal">
                            <span aria-hidden="true">&times;</span>
                            <span class="sr-only">Close</span>
                        </button>
                        <h4 class="modal-title" id="myModalLabel">Ny Bokning
                        </h4>
                    </div>
                    <!-- Modal Body -->
                    <div class="modal-body" id="modalBody">
                        <form role="form" id="fileForm">
                            <div class="form-group">
                                <label>BookningsID</label>
                                <input class="form-control" id="BookningsID" />
                            </div>
                            <div class="form-group">
                                <label>Spelare1</label>
                                <input class="form-control" id="Spelare1" />
                            </div>
                            <div class="form-group">
                                <label>Spelare2</label>
                                <input class="form-control" id="Spelare2" />
                            </div>
                            <div class="form-group">
                                <label>Spelare3</label>
                                <input class="form-control" id="Spelare3" />
                            </div>
                            <div class="form-group">
                                <label>Spelare4</label>
                                <input class="form-control" id="Spelare4" />
                            </div>
                            <div class="form-group">
                                <label>AntalSpelare</label>
                                <select name="players" form="players" id="players">
                                  <option value="1">1</option>
                                  <option value="2">2</option>
                                  <option value="3">3</option>
                                  <option value="4">4</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label>BokningsDate</label>
                                <input type="date" class="form-control" id="BokningsDate" />
                            </div>
                            <!-- hidden controls -->
                            <div style="display: none">
                                <input id="etag" />
                                <input id="bokningId" />
                            </div>
                        </form>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-danger" data-dismiss="modal" id="updateBokningLabel">
                                Avbryt
                            </button>
                            <button type="submit" class="btn btn-primary" id="bokningFormSubmit">
                                Skapa
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

</asp:Content>
