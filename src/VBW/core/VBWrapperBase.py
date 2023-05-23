from subprocess import PIPE, Popen
import os
from core.VBErrorDictionary import errorDictionary

class VB_ERROR(SyntaxError):
    pass

class INTERPRETER_PROCESS:
    def __init__(self, cScript, iScript, silenceExceptions):
        self.p = Popen([cScript,
                        '//nologo',
                        iScript],
                       stdout=PIPE,
                       stdin=PIPE,
                       encoding='ascii')
        self.errorStatus = 0
        self.silenceExceptions = silenceExceptions

    def communicate(self, message, printing=True):
        assert self.p is not None
        assert self.errorStatus == 0
        self.p.stdin.write(message + "\n")
        self.p.stdin.flush()
        answer = self.p.stdout.readline()
        if "\n" not in answer or answer == "!>!>!>ERROR\n":
            self.p.kill()
            self.errorStatus = 1 if "\n" not in answer else 2 
            if not self.silenceExceptions:
                if "\n" in answer:
                    errorNumber = self.p.stdout.readline()[:-1]
                    error = errorDictionary[errorNumber] if errorNumber in errorDictionary else errorNumber
                    errorMessage = f"VBS interpreter failed with error({error}) after receiving message: \"{message}\""
                else:
                    errorMessage = f"No response from VBS interpreter after message: \"{message}\""
                raise VB_ERROR(errorMessage)
            return "!>!>!>END\n"
        else:
            if printing:
                print(answer, end="")
            return answer

    def kill(self):
        self.p.kill()
        self.p = None

class VB_WRAPPER_BASE:

    # WRAPPER
    def __init__(self, path2CScript = None, path2InterpreterScript = None, silenceExceptions = True):
        self.cScript = path2CScript
        self.iScript = path2InterpreterScript
        self.silenceExceptions = silenceExceptions
        self.history = []
        self.interpreter = None
        # It goes to the default path of CScript.exe in the most common configurations
        if path2CScript is None:
            self.cScript = "C:/Windows/System32/CScript.exe"
        # By default it looks for the interactive interpreter inside the same directory as the main file
        if path2InterpreterScript is None:
            self.iScript = os.path.join(os.path.split(os.path.realpath(__file__))[0], "interactive_interpreter.vbs")

    # Interpreter Manipulation

    def createInterpreter(self,  startup_commands=None):
        if self.interpreter is not None:
            self.interpreter.kill()
        self.history = []
        self.interpreter = INTERPRETER_PROCESS(self.cScript, self.iScript, self.silenceExceptions)
        self.interpreterInitialization( startup_commands= [] if startup_commands is None else startup_commands)

    def communicateWithInterpreter(self, message, printing=True, recording=True):
        assert self.interpreter is not None
        if self.doRecordMessage(message,recording):
            self.history.append(message)
        return self.interpreter.communicate(message,printing)

    def doRecordMessage(self,message,recording):
        return recording

    def interpreterInitialization(self, startup_commands):
        assert self.interpreter is not None
        for sc in startup_commands:
            self.communicateWithInterpreter(sc, printing=False, recording=False)

    def destroyInterpreter(self):
        assert self.interpreter is not None
        self.cExit()
        self.interpreter.kill()
        self.history = []
        self.interpreter = None

    # Core Commands

    def cExec(self,message, printing=False, recording=True):
        message_sent = self.communicateWithInterpreter(message + "'x", printing= printing, recording= recording)
        return (message+"'x\n") == message_sent

    def cEval(self,message, printing=True, recording=True):
        answer = self.communicateWithInterpreter(message, printing= printing, recording= recording)[:-1]
        return answer

    def cExit(self):
        return self.communicateWithInterpreter("'e", printing=False, recording=False)

    # Recovery commands

    def cGetRecoveryCommandsArray(self):
        n = self.cEval("UBound(recoveryCommandsArray)", printing=False, recording=False)
        recoveryCommandsArray = [self.cEval(f"recoveryCommandsArray({n})", printing=False, recording=False) for n in range(int(n)+1)]
        return recoveryCommandsArray

    def cAddRecoveryCommand(self, recoveryCommand):
        recoveryCommandsArray = self.cGetRecoveryCommandsArray()
        try:
            index = recoveryCommandsArray.index("")
        except ValueError:
            index = len(recoveryCommandsArray)
            self.cExec(f"ReDim Preserve recoveryCommandsArray({index})")
        self.cExec(f"recoveryCommandsArray({index}) = \"{recoveryCommand}\"")

    def cAddRecoveryCommandFirst(self, recoveryCommand):
        recoveryCommandsArray = self.cGetRecoveryCommandsArray()
        index = len(recoveryCommandsArray)
        self.cExec(f"ReDim recoveryCommandsArray({index})")
        self.cExec(f"recoveryCommandsArray(1) = \"{recoveryCommand}\"")
        for x in range(len(recoveryCommandsArray)):
            self.cExec(f"recoveryCommandsArray({x+1}) = \"{recoveryCommandsArray[x]}\"")

    def cRemoveRecoveryCommand(self, recoveryCommand):
        recoveryCommandsArray = self.cGetRecoveryCommandsArray()
        if recoveryCommand in recoveryCommandsArray:
            index = recoveryCommandsArray.index(recoveryCommand)
            for i in range(index+1, len(recoveryCommandsArray)):
                self.cExec(f"recoveryCommandsArray({i-1}) = \"{recoveryCommandsArray[i]}\"")
            if len(recoveryCommandsArray) > 1:
                self.cExec(f"ReDim Preserve recoveryCommandsArray({len(recoveryCommandsArray)-2})")
            else:
                self.cExec(f"recoveryCommandsArray({index}) = \"\"")
            return True

        return False
