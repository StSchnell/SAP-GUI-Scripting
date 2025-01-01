var jacob = JavaImporter(
  com.jacob.activeX.ActiveXComponent,
  com.jacob.com.ComThread,
  com.jacob.com.Dispatch,
  com.jacob.com.Variant
);

function main() {

  const OUTPUT_CONSOLE = 0;
  const OUTPUT_WINDOW = 1;
  const OUTPUT_BUFFER = 2;

  with (jacob) {

    // Invokes SAP GUI Scripting method
    call = (session, id, name, parameter) => {
      if (typeof parameter === "undefined" || parameter === null) {
        return ActiveXComponent(session.invoke("findById", id).toDispatch())
          .invoke(name);
      } else {
        return ActiveXComponent(session.invoke("findById", id).toDispatch())
          .invoke(name, parameter);
      }
    }

    // Sets SAP GUI Scripting attribute
    set = (session, id, name, parameter) => {
      if (typeof parameter === "undefined" || parameter === null) {
        ActiveXComponent(session.invoke("findById", id).toDispatch())
          .setProperty(name);
      } else {
        ActiveXComponent(session.invoke("findById", id).toDispatch())
          .setProperty(name, parameter);
      }
    }

    // Gets SAP GUI Scripting attribute
    get = (session, id, name) => {
      return ActiveXComponent(session.invoke("findById", id).toDispatch())
        .getProperty(name);
    }

    ComThread.InitSTA();

    var SAPROTWr = new ActiveXComponent("SapROTWr.SapROTWrapper");
    var ROTEntry = SAPROTWr.invoke("GetROTEntry", "SAPGUI").toDispatch();

    try {

      var ScriptEngine = Dispatch.call(ROTEntry, "GetScriptingEngine");
      var application = new ActiveXComponent(ScriptEngine.toDispatch());

      application.setProperty("HistoryEnabled", false);

      var connection = new ActiveXComponent(
        application.invoke("Children", 0).toDispatch()
      );

      var DisabledByServer = connection.getProperty("DisabledByServer")
        .changeType(Variant.VariantBoolean).getBoolean();
      if(DisabledByServer == true) {
        print("Scripting is disabled by server");
        ComThread.Release();
        return;
      }

      var session = new ActiveXComponent(
        connection.invoke("Children", 0).toDispatch()
      );

      var Busy = session.getProperty("Busy")
        .changeType(Variant.VariantBoolean).getBoolean();
      if(Busy == true) {
        print("Session is busy");
        ComThread.Release();
        return;
      }

      var IsLowSpeedConnection = session.getPropertyAsComponent("Info")
        .getProperty("IsLowSpeedConnection")
        .changeType(Variant.VariantBoolean).getBoolean();
      if(IsLowSpeedConnection == true) {
        print("Connection is low speed");
        ComThread.Release();
        return;
      }

       set(session, "wnd[0]/tbar[0]/okcd", "text", "/nSE16");
      call(session, "wnd[0]", "sendVKey", 0);
       set(session, "wnd[0]/usr/ctxtDATABROWSE-TABLENAME", "text", "TADIR");
       set(session, "wnd[0]/usr/ctxtDATABROWSE-TABLENAME", "caretPosition", 5);
      call(session, "wnd[0]", "sendVKey", 0);
      call(session, "wnd[0]", "sendVKey", 31);
      call(session, "wnd[0]", "sendVKey", 0);
      call(session, "wnd[0]", "sendVKey", 3);
      call(session, "wnd[0]/tbar[0]/btn[3]", "press");

    } catch (exception) {
      print(exception);
    } finally {
      ComThread.Release();
    }

  }
  
}

main();
