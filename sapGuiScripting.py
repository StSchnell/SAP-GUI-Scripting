import clr
import platform
import System

clr.AddReference("Microsoft.VisualBasic")
import Microsoft.VisualBasic

def createObject(ProgId):
    """Creates and returns a reference to a COM object.
       @param {string} ProdId - Program ID of the object to create.
       @returns {object} - Reference to a COM object.
    """
    returnValue = None
    try:
        returnValue = Microsoft.VisualBasic.Interaction.CreateObject(ProgId)
    except Exception as ex:
        print("createObject exception: " + str(ex))
    return returnValue

def getObject(Class):
    """Returns a reference to an object provided by a COM component.
       @param {string} Class - Name representing the class of the object.
       @returns {__ComObject} - Reference to an object provided by a COM component.
    """
    returnValue = None
    try:
        returnValue = Microsoft.VisualBasic.Interaction.GetObject(Class)
    except Exception as ex:
        print("getObject exception: " + str(ex))
    return returnValue

def freeObject(Object):
    """Decrements the reference count of the specified COM object.
       @param {__ComObject} Object - The COM object to release.
    """
    try:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Object)
    except Exception as ex:
        print("freeObject exception: " + str(ex))

def getProperty(Object, PropertyName, PropertyParameter = None):
    """Gets a specific property.
       @param {__ComObject} Object - Object to get the property.
       @param {string} PropertyName - Name of the public property to get.
       @param {array.<object>} PropertyParameter - Arguments to pass.
       @returns {object} - Object representing the property.
    """
    returnValue = None
    try:
        returnValue = Object.GetType().InvokeMember(
            PropertyName,
            System.Reflection.BindingFlags.GetProperty,
            None,
            Object,
            PropertyParameter
        )
    except Exception as ex:
        print("getProperty exception: " + str(ex))
    return returnValue

def setProperty(Object, PropertyName, PropertyValue):
    """Sets a specific property.
       @param {__ComObject} Object - Object to set the property.
       @param {string} PropertyName - Name of the public property to set.
       @param {array.<object>} PropertyParameter - Arguments to pass.
    """
    try:
        Object.GetType().InvokeMember(
            PropertyName,
            System.Reflection.BindingFlags.SetProperty,
            None,
            Object,
            PropertyValue
        )
    except Exception as ex:
        print("setProperty exception: " + str(ex))

def invokeMethod(Object, MethodName, MethodParameter = None):
    """Invokes a specific method.
       @param {__ComObject} Object - Object to invoke the method.
       @param {string} MethodName - Name of the public method to invoke.
       @param {array.<object>} MethodParameter - Arguments to pass.
       @returns {object} - Object representing the result of the method.
    """
    returnValue = None
    try:
        returnValue = Object.GetType().InvokeMember(
            MethodName,
            System.Reflection.BindingFlags.InvokeMethod,
            None,
            Object,
            MethodParameter
        )
    except Exception as ex:
        print("invokeMethod exception: " + str(ex))
    return returnValue

def do(session, id, name, parameter = None):
    """Executes an SAP GUI Scripting action.
       @param {__ComObject} session - Session object to execute the action.
       @param {string} id - ID of the object to do the action.
       @param {string} name - Name of the method or property.
       @param {array.<object>} parameter - Arguments to pass.
    """
    try:
        invokeMethod(
            invokeMethod(session, "findById", [id]), name, parameter
        )
    except Exception as ex:
        print("do exception: " + str(ex))

def main():

    try:

        SapGuiAuto = getObject(Class = "SAPGUI")
        if SapGuiAuto == None:
            print("Can not get SapGuiAuto")
            return

        application = invokeMethod(SapGuiAuto, "GetScriptingEngine")
        if application == None:
            print("Can not get application")
            return

        setProperty(application, "HistoryEnabled", [False])

        connection = invokeMethod(
            getProperty(application, "Children"), "ElementAt", [0]
        )
        if connection == None:
            print("Can not get connection")
            return

        if getProperty(connection, "DisabledByServer") == True:
            print("Scripting is disabled by server")
            return

        session = invokeMethod(
            getProperty(connection, "Children"), "ElementAt", [0]
        )
        if session == None:
            print("Can not get session")
            return

        if getProperty(session, "Busy") == True:
            print("Session is busy")
            return
        info = getProperty(session, "Info")
        if getProperty(info, "IsLowSpeedConnection") == True:
            print("Connection is low speed")
            return

        do(session, "wnd[0]/tbar[0]/okcd", "text", ["/nse16"])
        do(session, "wnd[0]", "sendVKey", [0])
        do(session, "wnd[0]/usr/ctxtDATABROWSE-TABLENAME", "text", ["TADIR"])
        do(session, "wnd[0]/usr/ctxtDATABROWSE-TABLENAME", "caretPosition", [5])
        do(session, "wnd[0]", "sendVKey", [0])
        do(session, "wnd[0]/tbar[1]/btn[31]", "press")
        do(session, "wnd[1]/tbar[0]/btn[0]", "press")
        do(session, "wnd[0]/tbar[0]/btn[3]", "press")
        do(session, "wnd[0]", "sendVKey", [3])

    except Exception as ex:
        print("Exception: " + ex + "\n" + ex.__cause__)

    finally:
        setProperty(application, "HistoryEnabled", [False])

if __name__ == '__main__':
    if platform.system() == "Windows":
        main()
