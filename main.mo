import OrderedMap "mo:base/OrderedMap";
import Text "mo:base/Text";
import Nat "mo:base/Nat";
import Time "mo:base/Time";
import Iter "mo:base/Iter";
import List "mo:base/List";
import Debug "mo:base/Debug";
import Int "mo:base/Int";
import Principal "mo:base/Principal";
import AccessControl "authorization/access-control";
actor Spreadsheet {
  transient let textMap = OrderedMap.Make<Text>(Text.compare);
  transient let principalMap = OrderedMap.Make<Principal>(Principal.compare);
  type CellFormat = {
    bold : ?Bool;
    italic : ?Bool;
    underline : ?Bool;
    fontFamily : ?Text;
    fontSize : ?Nat;
    fontColor : ?Text;
    fillColor : ?Text;
    alignment : ?Text;
    borders : ?Text;
  };
  type Cell = {
    value : Text;
    formula : ?Text;
    format : ?CellFormat;
  };
  type ImageData = {
    id : Text;
    src : Text;
    x : Int;
    y : Int;
    width : Int;
    height : Int;
    anchorCell : Text;
  };
  type Sheet = {
    name : Text;
    cells : OrderedMap.Map<Text, Cell>;
    images : OrderedMap.Map<Text, ImageData>;
  };
  type SpreadsheetPermission = {
    #owner;
    #editor;
    #viewer;
  };
  type SpreadsheetFile = {
    id : Text;
    name : Text;
    owner : Principal;
    createdAt : Int;
    sheets : OrderedMap.Map<Text, Sheet>;
    permissions : OrderedMap.Map<Principal, SpreadsheetPermission>;
    activeSheet : Text;
  };
  type UserProfile = {
    name : Text;
  };
  var spreadsheets : OrderedMap.Map<Text, SpreadsheetFile> = textMap.empty();
  var userProfiles : OrderedMap.Map<Principal, UserProfile> = principalMap.empty();
  let accessControlState = AccessControl.initState();
  private func hasSpreadsheetPermission(spreadsheet : SpreadsheetFile, caller : Principal, requiredPermission : SpreadsheetPermission) : Bool {
    if (Principal.equal(spreadsheet.owner, caller)) {
      return true;
    };
    switch (principalMap.get(spreadsheet.permissions, caller)) {
      case (null) { false };
      case (?permission) {
        switch (requiredPermission) {
          case (#owner) { false };
          case (#editor) {
            switch (permission) {
              case (#owner) { true };
              case (#editor) { true };
              case (#viewer) { false };
            };
          };
          case (#viewer) { true };
        };
      };
    };
  };
  public shared ({ caller }) func initializeAccessControl() : async () {
    AccessControl.initialize(accessControlState, caller);
  };
  public query ({ caller }) func getCallerUserRole() : async AccessControl.UserRole {
    AccessControl.getUserRole(accessControlState, caller);
  };
  public shared ({ caller }) func assignCallerUserRole(user : Principal, role : AccessControl.UserRole) : async () {
    AccessControl.assignRole(accessControlState, caller, user, role);
  };
  public query ({ caller }) func isCallerAdmin() : async Bool {
    AccessControl.isAdmin(accessControlState, caller);
  };
  public query ({ caller }) func getCallerUserProfile() : async ?UserProfile {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access profiles");
    };
    principalMap.get(userProfiles, caller);
  };
  public query ({ caller }) func getUserProfile(user : Principal) : async ?UserProfile {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access profiles");
    };
    if (caller != user and not AccessControl.isAdmin(accessControlState, caller)) {
      Debug.trap("Unauthorized: Can only view your own profile");
    };
    principalMap.get(userProfiles, user);
  };
  public shared ({ caller }) func saveCallerUserProfile(profile : UserProfile) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can save profiles");
    };
    userProfiles := principalMap.put(userProfiles, caller, profile);
  };
  public shared ({ caller }) func createSpreadsheet(name : Text) : async Text {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can create spreadsheets");
    };
    let id = Text.concat(name, Int.toText(Time.now()));
    let defaultSheet : Sheet = {
      name = "Sheet1";
      cells = textMap.empty();
      images = textMap.empty();
    };
    let sheets = textMap.put(textMap.empty(), "Sheet1", defaultSheet);
    let spreadsheet : SpreadsheetFile = {
      id;
      name;
      owner = caller;
      createdAt = Time.now();
      sheets;
      permissions = principalMap.empty();
      activeSheet = "Sheet1";
    };
    spreadsheets := textMap.put(spreadsheets, id, spreadsheet);
    id;
  };
  public query ({ caller }) func getSpreadsheet(id : Text) : async ?SpreadsheetFile {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access spreadsheets");
    };
    switch (textMap.get(spreadsheets, id)) {
      case (null) { null };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to view this spreadsheet");
        };
        ?spreadsheet;
      };
    };
  };
  public shared ({ caller }) func saveCell(spreadsheetId : Text, sheetName : Text, cellId : Text, value : Text, formula : ?Text) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can modify spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to edit this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { Debug.trap("Sheet not found") };
          case (?sheet) {
            let cell : Cell = {
              value;
              formula;
              format = null;
            };
            let updatedCells = textMap.put(sheet.cells, cellId, cell);
            let updatedSheet : Sheet = {
              name = sheet.name;
              cells = updatedCells;
              images = sheet.images;
            };
            let updatedSheets = textMap.put(spreadsheet.sheets, sheetName, updatedSheet);
            let updatedSpreadsheet : SpreadsheetFile = {
              id = spreadsheet.id;
              name = spreadsheet.name;
              owner = spreadsheet.owner;
              createdAt = spreadsheet.createdAt;
              sheets = updatedSheets;
              permissions = spreadsheet.permissions;
              activeSheet = spreadsheet.activeSheet;
            };
            spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
          };
        };
      };
    };
  };
  public shared ({ caller }) func addSheet(spreadsheetId : Text) : async Text {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can modify spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to edit this spreadsheet");
        };
        var sheetNumber = 2;
        var sheetName = Text.concat("Sheet", Nat.toText(sheetNumber));
        while (textMap.get(spreadsheet.sheets, sheetName) != null) {
          sheetNumber += 1;
          sheetName := Text.concat("Sheet", Nat.toText(sheetNumber));
        };
        let newSheet : Sheet = {
          name = sheetName;
          cells = textMap.empty();
          images = textMap.empty();
        };
        let updatedSheets = textMap.put(spreadsheet.sheets, sheetName, newSheet);
        let updatedSpreadsheet : SpreadsheetFile = {
          id = spreadsheet.id;
          name = spreadsheet.name;
          owner = spreadsheet.owner;
          createdAt = spreadsheet.createdAt;
          sheets = updatedSheets;
          permissions = spreadsheet.permissions;
          activeSheet = spreadsheet.activeSheet;
        };
        spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
        sheetName;
      };
    };
  };
  public shared ({ caller }) func switchSheet(spreadsheetId : Text, sheetName : Text) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can switch sheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to switch sheets in this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { Debug.trap("Sheet not found") };
          case (?_) {
            let updatedSpreadsheet : SpreadsheetFile = {
              id = spreadsheet.id;
              name = spreadsheet.name;
              owner = spreadsheet.owner;
              createdAt = spreadsheet.createdAt;
              sheets = spreadsheet.sheets;
              permissions = spreadsheet.permissions;
              activeSheet = sheetName;
            };
            spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
          };
        };
      };
    };
  };
  public query ({ caller }) func getActiveSheet(spreadsheetId : Text) : async ?Text {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access active sheet");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { null };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to view active sheet in this spreadsheet");
        };
        ?spreadsheet.activeSheet;
      };
    };
  };
  public query ({ caller }) func getCell(spreadsheetId : Text, sheetName : Text, cellId : Text) : async ?Cell {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { null };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to view this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { null };
          case (?sheet) {
            textMap.get(sheet.cells, cellId);
          };
        };
      };
    };
  };
  public query ({ caller }) func listSpreadsheets() : async [SpreadsheetFile] {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can list spreadsheets");
    };
    let allSpreadsheets = Iter.toArray(textMap.vals(spreadsheets));
    var accessibleSpreadsheets = List.nil<SpreadsheetFile>();
    for (spreadsheet in allSpreadsheets.vals()) {
      if (hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
        accessibleSpreadsheets := List.push(spreadsheet, accessibleSpreadsheets);
      };
    };
    List.toArray(accessibleSpreadsheets);
  };
  public query ({ caller }) func listSheets(spreadsheetId : Text) : async [Text] {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { [] };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to view this spreadsheet");
        };
        Iter.toArray(textMap.keys(spreadsheet.sheets));
      };
    };
  };
  public query ({ caller }) func getSheet(spreadsheetId : Text, sheetName : Text) : async ?Sheet {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { null };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to view this spreadsheet");
        };
        textMap.get(spreadsheet.sheets, sheetName);
      };
    };
  };
  public query ({ caller }) func getRange(spreadsheetId : Text, sheetName : Text, cellIds : [Text]) : async [Cell] {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { [] };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to view this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { [] };
          case (?sheet) {
            var cells = List.nil<Cell>();
            for (cellId in cellIds.vals()) {
              switch (textMap.get(sheet.cells, cellId)) {
                case (null) {};
                case (?cell) {
                  cells := List.push(cell, cells);
                };
              };
            };
            List.toArray(cells);
          };
        };
      };
    };
  };
  public shared ({ caller }) func shareSpreadsheet(spreadsheetId : Text, user : Principal, permission : SpreadsheetPermission) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can share spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #owner)) {
          Debug.trap("Unauthorized: Only owner can share spreadsheet");
        };
        let updatedPermissions = principalMap.put(spreadsheet.permissions, user, permission);
        let updatedSpreadsheet : SpreadsheetFile = {
          id = spreadsheet.id;
          name = spreadsheet.name;
          owner = spreadsheet.owner;
          createdAt = spreadsheet.createdAt;
          sheets = spreadsheet.sheets;
          permissions = updatedPermissions;
          activeSheet = spreadsheet.activeSheet;
        };
        spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
      };
    };
  };
  public shared ({ caller }) func revokeSpreadsheetAccess(spreadsheetId : Text, user : Principal) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can revoke access");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #owner)) {
          Debug.trap("Unauthorized: Only owner can revoke access");
        };
        let updatedPermissions = principalMap.delete(spreadsheet.permissions, user);
        let updatedSpreadsheet : SpreadsheetFile = {
          id = spreadsheet.id;
          name = spreadsheet.name;
          owner = spreadsheet.owner;
          createdAt = spreadsheet.createdAt;
          sheets = spreadsheet.sheets;
          permissions = updatedPermissions;
          activeSheet = spreadsheet.activeSheet;
        };
        spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
      };
    };
  };
  public shared ({ caller }) func deleteSpreadsheet(spreadsheetId : Text) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can delete spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #owner)) {
          Debug.trap("Unauthorized: Only owner can delete spreadsheet");
        };
        spreadsheets := textMap.delete(spreadsheets, spreadsheetId);
      };
    };
  };
  public shared ({ caller }) func deleteSheet(spreadsheetId : Text, sheetName : Text) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can modify spreadsheets");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to edit this spreadsheet");
        };
        let updatedSheets = textMap.delete(spreadsheet.sheets, sheetName);
        let updatedSpreadsheet : SpreadsheetFile = {
          id = spreadsheet.id;
          name = spreadsheet.name;
          owner = spreadsheet.owner;
          createdAt = spreadsheet.createdAt;
          sheets = updatedSheets;
          permissions = spreadsheet.permissions;
          activeSheet = spreadsheet.activeSheet;
        };
        spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
      };
    };
  };
  public shared ({ caller }) func applyFormatToSelection(spreadsheetId : Text, sheetName : Text, cellIds : [Text], formatObj : CellFormat) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can format cells");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to format this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { Debug.trap("Sheet not found") };
          case (?sheet) {
            var updatedCells = sheet.cells;
            for (cellId in cellIds.vals()) {
              switch (textMap.get(sheet.cells, cellId)) {
                case (null) {};
                case (?cell) {
                  let updatedCell : Cell = {
                    value = cell.value;
                    formula = cell.formula;
                    format = ?formatObj;
                  };
                  updatedCells := textMap.put(updatedCells, cellId, updatedCell);
                };
              };
            };
            let updatedSheet : Sheet = {
              name = sheet.name;
              cells = updatedCells;
              images = sheet.images;
            };
            let updatedSheets = textMap.put(spreadsheet.sheets, sheetName, updatedSheet);
            let updatedSpreadsheet : SpreadsheetFile = {
              id = spreadsheet.id;
              name = spreadsheet.name;
              owner = spreadsheet.owner;
              createdAt = spreadsheet.createdAt;
              sheets = updatedSheets;
              permissions = spreadsheet.permissions;
              activeSheet = spreadsheet.activeSheet;
            };
            spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
          };
        };
      };
    };
  };
  public shared ({ caller }) func applyFontColor(spreadsheetId : Text, sheetName : Text, cellIds : [Text], color : Text) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can apply font color");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to apply font color to this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { Debug.trap("Sheet not found") };
          case (?sheet) {
            var updatedCells = sheet.cells;
            for (cellId in cellIds.vals()) {
              switch (textMap.get(sheet.cells, cellId)) {
                case (null) {};
                case (?cell) {
                  let currentFormat = switch (cell.format) {
                    case (null) { { bold = null; italic = null; underline = null; fontFamily = null; fontSize = null; fontColor = null; fillColor = null; alignment = null; borders = null } };
                    case (?format) { format };
                  };
                  let updatedFormat = { currentFormat with fontColor = ?color };
                  let updatedCell : Cell = {
                    value = cell.value;
                    formula = cell.formula;
                    format = ?updatedFormat;
                  };
                  updatedCells := textMap.put(updatedCells, cellId, updatedCell);
                };
              };
            };
            let updatedSheet : Sheet = {
              name = sheet.name;
              cells = updatedCells;
              images = sheet.images;
            };
            let updatedSheets = textMap.put(spreadsheet.sheets, sheetName, updatedSheet);
            let updatedSpreadsheet : SpreadsheetFile = {
              id = spreadsheet.id;
              name = spreadsheet.name;
              owner = spreadsheet.owner;
              createdAt = spreadsheet.createdAt;
              sheets = updatedSheets;
              permissions = spreadsheet.permissions;
              activeSheet = spreadsheet.activeSheet;
            };
            spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
          };
        };
      };
    };
  };
  public shared ({ caller }) func applyFillColor(spreadsheetId : Text, sheetName : Text, cellIds : [Text], color : Text) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can apply fill color");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to apply fill color to this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { Debug.trap("Sheet not found") };
          case (?sheet) {
            var updatedCells = sheet.cells;
            for (cellId in cellIds.vals()) {
              switch (textMap.get(sheet.cells, cellId)) {
                case (null) {};
                case (?cell) {
                  let currentFormat = switch (cell.format) {
                    case (null) { { bold = null; italic = null; underline = null; fontFamily = null; fontSize = null; fontColor = null; fillColor = null; alignment = null; borders = null } };
                    case (?format) { format };
                  };
                  let updatedFormat = { currentFormat with fillColor = ?color };
                  let updatedCell : Cell = {
                    value = cell.value;
                    formula = cell.formula;
                    format = ?updatedFormat;
                  };
                  updatedCells := textMap.put(updatedCells, cellId, updatedCell);
                };
              };
            };
            let updatedSheet : Sheet = {
              name = sheet.name;
              cells = updatedCells;
              images = sheet.images;
            };
            let updatedSheets = textMap.put(spreadsheet.sheets, sheetName, updatedSheet);
            let updatedSpreadsheet : SpreadsheetFile = {
              id = spreadsheet.id;
              name = spreadsheet.name;
              owner = spreadsheet.owner;
              createdAt = spreadsheet.createdAt;
              sheets = updatedSheets;
              permissions = spreadsheet.permissions;
              activeSheet = spreadsheet.activeSheet;
            };
            spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
          };
        };
      };
    };
  };
  public shared ({ caller }) func addImage(spreadsheetId : Text, sheetName : Text, image : ImageData) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can add images");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to add images to this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { Debug.trap("Sheet not found") };
          case (?sheet) {
            let updatedImages = textMap.put(sheet.images, image.id, image);
            let updatedSheet : Sheet = {
              name = sheet.name;
              cells = sheet.cells;
              images = updatedImages;
            };
            let updatedSheets = textMap.put(spreadsheet.sheets, sheetName, updatedSheet);
            let updatedSpreadsheet : SpreadsheetFile = {
              id = spreadsheet.id;
              name = spreadsheet.name;
              owner = spreadsheet.owner;
              createdAt = spreadsheet.createdAt;
              sheets = updatedSheets;
              permissions = spreadsheet.permissions;
              activeSheet = spreadsheet.activeSheet;
            };
            spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
          };
        };
      };
    };
  };
  public shared ({ caller }) func updateImage(spreadsheetId : Text, sheetName : Text, image : ImageData) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can update images");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to update images in this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { Debug.trap("Sheet not found") };
          case (?sheet) {
            let updatedImages = textMap.put(sheet.images, image.id, image);
            let updatedSheet : Sheet = {
              name = sheet.name;
              cells = sheet.cells;
              images = updatedImages;
            };
            let updatedSheets = textMap.put(spreadsheet.sheets, sheetName, updatedSheet);
            let updatedSpreadsheet : SpreadsheetFile = {
              id = spreadsheet.id;
              name = spreadsheet.name;
              owner = spreadsheet.owner;
              createdAt = spreadsheet.createdAt;
              sheets = updatedSheets;
              permissions = spreadsheet.permissions;
              activeSheet = spreadsheet.activeSheet;
            };
            spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
          };
        };
      };
    };
  };
  public shared ({ caller }) func deleteImage(spreadsheetId : Text, sheetName : Text, imageId : Text) : async () {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can delete images");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { Debug.trap("Spreadsheet not found") };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #editor)) {
          Debug.trap("Unauthorized: No permission to delete images from this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { Debug.trap("Sheet not found") };
          case (?sheet) {
            let updatedImages = textMap.delete(sheet.images, imageId);
            let updatedSheet : Sheet = {
              name = sheet.name;
              cells = sheet.cells;
              images = updatedImages;
            };
            let updatedSheets = textMap.put(spreadsheet.sheets, sheetName, updatedSheet);
            let updatedSpreadsheet : SpreadsheetFile = {
              id = spreadsheet.id;
              name = spreadsheet.name;
              owner = spreadsheet.owner;
              createdAt = spreadsheet.createdAt;
              sheets = updatedSheets;
              permissions = spreadsheet.permissions;
              activeSheet = spreadsheet.activeSheet;
            };
            spreadsheets := textMap.put(spreadsheets, spreadsheetId, updatedSpreadsheet);
          };
        };
      };
    };
  };
  public query ({ caller }) func getImage(spreadsheetId : Text, sheetName : Text, imageId : Text) : async ?ImageData {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access images");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { null };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to view images in this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { null };
          case (?sheet) {
            textMap.get(sheet.images, imageId);
          };
        };
      };
    };
  };
  public query ({ caller }) func listImages(spreadsheetId : Text, sheetName : Text) : async [ImageData] {
    if (not (AccessControl.hasPermission(accessControlState, caller, #user))) {
      Debug.trap("Unauthorized: Only users can access images");
    };
    switch (textMap.get(spreadsheets, spreadsheetId)) {
      case (null) { [] };
      case (?spreadsheet) {
        if (not hasSpreadsheetPermission(spreadsheet, caller, #viewer)) {
          Debug.trap("Unauthorized: No permission to view images in this spreadsheet");
        };
        switch (textMap.get(spreadsheet.sheets, sheetName)) {
          case (null) { [] };
          case (?sheet) {
            Iter.toArray(textMap.vals(sheet.images));
          };
        };
      };
    };
  };
};

