from core.VBInterpreterProcess import INTERPRETER_PROCESS


class VB_WRAPPER_BASE:

    # WRAPPER

    def __init__(self, path2CScript=None, path2InterpreterScript=None,
                 silenceExceptions=True, default_startup_commands=None):
        self.cScript = path2CScript
        self.iScript = path2InterpreterScript
        self.silenceExceptions = silenceExceptions
        self.history = []
        self.interpreter = None
        self.default_startup_commands = default_startup_commands

    def has_healthy_interpreter(self):
        return self.interpreter is not None and self.interpreter.healthy()

    # Interpreter Manipulation

    def createInterpreter(self,  startup_commands=None):
        if self.interpreter is not None:
            self.interpreter.kill()
        if startup_commands is None:
            startup_commands = self.default_startup_commands
        self.history = []
        self.interpreter = INTERPRETER_PROCESS(self.cScript, self.iScript, self.silenceExceptions)
        self.interpreterInitialization(startup_commands=[] if startup_commands is None else startup_commands)

    def communicateWithInterpreter(self, message, printing=True, recording=True):
        assert self.has_healthy_interpreter()
        if self.doRecordMessage(message,recording):
            self.history.append(message)
        return self.interpreter.communicate(message,printing)

    def doRecordMessage(self,message,recording):
        return recording

    def interpreterInitialization(self, startup_commands):
        assert self.has_healthy_interpreter()
        for sc in startup_commands:
            self.communicateWithInterpreter(sc, printing=False, recording=False)

    def destroyInterpreter(self):
        assert self.has_healthy_interpreter()
        self.cExit()
        self.interpreter.kill()
        self.history = []
        self.interpreter = None

    # Core Commands

    def cExec(self,message, printing=False, recording=True):
        assert self.has_healthy_interpreter()
        message_sent = self.communicateWithInterpreter(message + "'x", printing= printing, recording= recording)
        return (message+"'x\n") == message_sent

    def cEval(self,message, printing=True, recording=True):
        assert self.has_healthy_interpreter()
        answer = self.communicateWithInterpreter(message, printing= printing, recording= recording)[:-1]
        return answer

    def cExit(self):
        assert self.has_healthy_interpreter()
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

    # Context managers

    def __enter__(self):
        if not self.has_healthy_interpreter():
            self.createInterpreter()
        return self

    def __exit__(self, _type, _value, _traceback):
        if self.has_healthy_interpreter():
            self.destroyInterpreter()
