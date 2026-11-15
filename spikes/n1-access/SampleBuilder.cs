using System.Runtime.Versioning;

namespace AccessSpike;

/// <summary>
/// Builds the sample database through Access, from the text in sample/ and the forms laid out here: a start-up form
/// and an AutoExec macro that each leave a mark; a form with a subform that has a subform of its own; a continuous
/// form; a tab control, an option group and a combo box with a row source; events run by event procedures, by a
/// macro and by an expression; a standard module and a class module; local tables and a linked one. All of its
/// content is made up.
/// </summary>
[SupportedOSPlatform("windows")]
public static class SampleBuilder
{
  // Access's control types (acLabel and so on).
  private const int Label = 100, Button = 104, OptionButton = 105, OptionGroup = 107, TextBox = 109, ComboBox = 111, Subform = 112, TabControl = 123;
  private const int DbText = 10;

  /// <summary>
  /// Where the linked table's back end goes: a folder every Windows machine has, so the link's path names no person.
  /// </summary>
  public static string BackEndPath =>
    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "AccessDocSample", "SampleBackEnd.accdb");

  public static void Build(string sampleFolder, string databasePath)
  {
    databasePath = Path.GetFullPath(databasePath);
    Directory.CreateDirectory(Path.GetDirectoryName(databasePath)!);
    File.Delete(databasePath);
    Directory.CreateDirectory(Path.GetDirectoryName(BackEndPath)!);
    File.Delete(BackEndPath);

    AccessApp.Use(application => BuildWith(application, sampleFolder, databasePath));
  }

  private static void BuildWith(object application, string sampleFolder, string databasePath)
  {
    dynamic app = application; // application itself isn't dynamic, so a call that takes a lambda is an ordinary one

    // The back end first: one table for the front end to link to.
    var backEnd = app.DBEngine.CreateDatabase(BackEndPath, ";LANGID=0x0409;CP=1252;COUNTRY=0");
    backEnd.Execute("CREATE TABLE Regions (RegionName TEXT(20) CONSTRAINT PK_Regions PRIMARY KEY, Manager TEXT(50))");
    foreach (var region in new[] { "North", "South", "East", "West" })
    {
      backEnd.Execute($"INSERT INTO Regions (RegionName, Manager) VALUES ('{region}', 'Manager of {region}')");
    }

    backEnd.Close();

    app.NewCurrentDatabase(databasePath);
    var db = app.CurrentDb();

    // Access writes the Windows user's name into a new database as its Author, and their organisation as its
    // Company. The repository is public, so both are replaced before anything else is done.
    var summary = db.Containers["Databases"].Documents["SummaryInfo"];
    foreach (var (property, value) in new[] { ("Author", "Fixture Author"), ("Company", "Fixture Company"), ("Title", "accessdoc sample") })
    {
      try
      {
        summary.Properties[property].Value = value;
      }
      catch (Exception)
      {
        // Access didn't write that one here.
      }
    }

    db.Execute("CREATE TABLE Customers (Id AUTOINCREMENT CONSTRAINT PK_Customers PRIMARY KEY, CustomerName TEXT(80), Region TEXT(20), Standing LONG, Notes MEMO)");
    db.Execute("CREATE TABLE Orders (Id AUTOINCREMENT CONSTRAINT PK_Orders PRIMARY KEY, CustomerId LONG, OrderDate DATETIME, Status TEXT(10), Total CURRENCY)");
    db.Execute("CREATE TABLE OrderLines (Id AUTOINCREMENT CONSTRAINT PK_OrderLines PRIMARY KEY, OrderId LONG, Product TEXT(50), Quantity LONG, Price CURRENCY)");
    db.Execute("CREATE TABLE StartupLog (Id AUTOINCREMENT CONSTRAINT PK_StartupLog PRIMARY KEY, What TEXT(50), At DATETIME)");
    db.Execute("INSERT INTO Customers (CustomerName, Region, Standing) VALUES ('Acme Pty Ltd', 'North', 1)");
    db.Execute("INSERT INTO Customers (CustomerName, Region, Standing) VALUES ('Bluegum Traders', 'South', 2)");
    db.Execute("INSERT INTO Orders (CustomerId, OrderDate, Status, Total) VALUES (1, #2026-01-15#, 'NEW', 120)");
    db.Execute("INSERT INTO Orders (CustomerId, OrderDate, Status, Total) VALUES (1, #2026-02-03#, 'SENT', 45.5)");
    db.Execute("INSERT INTO OrderLines (OrderId, Product, Quantity, Price) VALUES (1, 'Widget', 4, 30)");
    db.Execute("INSERT INTO OrderLines (OrderId, Product, Quantity, Price) VALUES (2, 'Gadget', 1, 45.5)");

    Step("linked table and queries");
    var linked = db.CreateTableDef("Regions");
    linked.Connect = ";DATABASE=" + BackEndPath;
    linked.SourceTableName = "Regions";
    db.TableDefs.Append(linked);

    db.CreateQueryDef("qryOpenOrders", "SELECT Orders.Id, Orders.CustomerId, Orders.OrderDate, Orders.Status, Orders.Total FROM Orders WHERE Orders.Status <> 'SENT' ORDER BY Orders.OrderDate;");
    db.CreateQueryDef("qryCustomerTotals", "SELECT Customers.CustomerName, Sum(Orders.Total) AS Total FROM Customers INNER JOIN Orders ON Customers.Id = Orders.CustomerId GROUP BY Customers.CustomerName;");

    Step("modules");
    app.LoadFromText(AccessApp.AcModule, "basUtility", Path.Combine(sampleFolder, "basUtility.bas"));
    app.LoadFromText(AccessApp.AcModule, "clsOrderTotals", Path.Combine(sampleFolder, "clsOrderTotals.cls"));
    Step("macro mcrGreet");
    app.LoadFromText(AccessApp.AcMacro, "mcrGreet", Path.Combine(sampleFolder, "mcrGreet.macro"));

    // Innermost first, so each subform control has its form to hold.
    Form(application, sampleFolder, "frmOrderLines", "Order lines", "OrderLines", continuous: true, form =>
    {
      Bound(app, form, TextBox, "txtProduct", "Product", 200, 100, 2400);
      Bound(app, form, TextBox, "txtQuantity", "Quantity", 2700, 100, 900).AfterUpdate = "[Event Procedure]";
      Bound(app, form, TextBox, "txtPrice", "Price", 3700, 100, 1100);
      Control(app, form, TextBox, "txtLineTotal", 4900, 100, 1200).ControlSource = "=[Quantity]*[Price]";
    });

    Form(application, sampleFolder, "frmOrders", "Orders", "Orders", continuous: false, form =>
    {
      Bound(app, form, TextBox, "txtOrderId", "Id", 1500, 200, 1000);
      Bound(app, form, TextBox, "txtOrderDate", "OrderDate", 1500, 600, 1600);
      Bound(app, form, TextBox, "txtStatus", "Status", 1500, 1000, 1200);
      Bound(app, form, TextBox, "txtTotal", "Total", 1500, 1400, 1400);
      var button = Control(app, form, Button, "cmdLines", 3600, 200, 1500);
      button.Caption = "Lines…";
      button.OnClick = "[Event Procedure]";
      var lines = Control(app, form, Subform, "sfrLines", 200, 2000, 6600, 2200);
      lines.SourceObject = "Form.frmOrderLines";
      lines.LinkMasterFields = "Id";
      lines.LinkChildFields = "OrderId";
      form.OnCurrent = "[Event Procedure]";
    });

    Form(application, sampleFolder, "frmCustomers", "Customers", "Customers", continuous: false, form =>
    {
      Bound(app, form, TextBox, "txtCustomerName", "CustomerName", 1700, 200, 3000);
      var region = Bound(app, form, ComboBox, "cboRegion", "Region", 1700, 650, 2000);
      region.RowSourceType = "Table/Query";
      region.RowSource = "SELECT Regions.RegionName FROM Regions ORDER BY Regions.RegionName;";
      region.AfterUpdate = "[Event Procedure]";

      var standing = Bound(app, form, OptionGroup, "fraStanding", "Standing", 5200, 150, 2200, 1000);
      var good = Control(app, form, OptionButton, "optGood", 5400, 450, 300, 240, parent: "fraStanding");
      good.OptionValue = 1;
      var poor = Control(app, form, OptionButton, "optOnHold", 5400, 780, 300, 240, parent: "fraStanding");
      poor.OptionValue = 2;
      _ = standing;

      // A tab control comes with two pages: the first holds the orders, in a subform, and the second the notes.
      var tabs = Control(app, form, TabControl, "tabMain", 200, 1400, 7400, 5600);
      tabs.Pages[0].Name = "pgOrders";
      tabs.Pages[0].Caption = "Orders";
      tabs.Pages[1].Name = "pgNotes";
      tabs.Pages[1].Caption = "Notes";
      var orders = Control(app, form, Subform, "sfrOrders", 350, 1900, 7100, 4900, parent: "pgOrders");
      orders.SourceObject = "Form.frmOrders";
      orders.LinkMasterFields = "Id";
      orders.LinkChildFields = "CustomerId";
      Bound(app, form, TextBox, "txtNotes", "Notes", 350, 1900, 7000, 2000, parent: "pgNotes");

      var greet = Control(app, form, Button, "cmdGreet", 200, 7200, 1600);
      greet.Caption = "Greet (macro)";
      greet.OnClick = "mcrGreet";
      var hello = Control(app, form, Button, "cmdHello", 2000, 7200, 1900);
      hello.Caption = "Hello (expression)";
      hello.OnClick = "=SayHello()";
      var close = Control(app, form, Button, "cmdClose", 4100, 7200, 1200);
      close.Caption = "Close";
      close.OnClick = "[Event Procedure]";
      form.OnCurrent = "[Event Procedure]";
    });

    Form(application, sampleFolder, "frmStart", "Sample database", recordSource: null, continuous: false, form =>
    {
      var title = Control(app, form, Label, "lblTitle", 300, 300, 5000, 400);
      title.Caption = "A sample database for accessdoc";
      var button = Control(app, form, Button, "cmdCustomers", 300, 1000, 2000);
      button.Caption = "Customers";
      button.OnClick = "[Event Procedure]";
      form.OnLoad = "[Event Procedure]";
    });

    // Last, what runs when the database is opened: the start-up form, and the AutoExec macro.
    Step("macro AutoExec and start-up");
    app.LoadFromText(AccessApp.AcMacro, "AutoExec", Path.Combine(sampleFolder, "AutoExec.macro"));
    db.Properties.Append(db.CreateProperty("StartUpForm", DbText, "frmStart"));
    db.Properties.Append(db.CreateProperty("AppTitle", DbText, "accessdoc sample"));
    app.CloseCurrentDatabase();
  }

  private static void Step(string what) => Console.WriteLine("  " + what);

  private static void Form(object application, string sampleFolder, string name, string caption, string? recordSource, bool continuous, Action<dynamic> fill)
  {
    dynamic app = application;
    Step("form " + name);
    var form = app.CreateForm();
    string made = form.Name;
    form.Caption = caption;
    if (recordSource is not null)
    {
      form.RecordSource = recordSource;
    }

    form.DefaultView = continuous ? 1 : 0;
    fill(form);

    var code = Path.Combine(sampleFolder, name + ".vba");
    if (File.Exists(code))
    {
      form.HasModule = true;
      var module = form.Module;
      module.DeleteLines(1, module.CountOfLines);
      module.AddFromString(File.ReadAllText(code));
    }

    app.DoCmd.Close(AccessApp.AcForm, made, AccessApp.AcSaveYes);
    app.DoCmd.Rename(name, AccessApp.AcForm, made);
  }

  private static dynamic Control(dynamic app, dynamic form, int type, string name, int left, int top, int width, int height = 300, string? parent = null)
  {
    var control = app.CreateControl(form.Name, type, 0, parent ?? "", "", left, top, width, height);
    control.Name = name;
    return control;
  }

  private static dynamic Bound(dynamic app, dynamic form, int type, string name, string field, int left, int top, int width, int height = 300, string? parent = null)
  {
    var control = app.CreateControl(form.Name, type, 0, parent ?? "", field, left, top, width, height);
    control.Name = name;
    return control;
  }
}
