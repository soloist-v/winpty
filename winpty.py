import os
import ctypes
import _winapi
from ctypes import wintypes
from threading import Thread
from msvcrt import open_osfhandle

WINPTY_SPAWN_FLAG_AUTO_SHUTDOWN = 1
# open stderr pipe
WINPTY_FLAG_CONERR = 0x1
# disable color output
WINPTY_FLAG_PLAIN_OUTPUT = 0x2
# enable color output (the default is enable)
WINPTY_FLAG_COLOR_ESCAPES = 0x4
WINPTY_FLAG_ALLOW_CURPROC_DESKTOP_CREATION = 0x8  #
win_pty_dll = ctypes.windll.LoadLibrary("winpty.dll")
# Error handling...
winpty_error_code = win_pty_dll.winpty_error_code
winpty_error_msg = win_pty_dll.winpty_error_msg
winpty_error_free = win_pty_dll.winpty_error_free
# Configuration of a new agent.
winpty_config_new = win_pty_dll.winpty_config_new
winpty_config_free = win_pty_dll.winpty_config_free
winpty_config_set_initial_size = win_pty_dll.winpty_config_set_initial_size
winpty_config_set_mouse_mode = win_pty_dll.winpty_config_set_mouse_mode
winpty_config_set_agent_timeout = win_pty_dll.winpty_config_set_agent_timeout
# Start the agent.
winpty_open = win_pty_dll.winpty_open
winpty_agent_process = win_pty_dll.winpty_agent_process
# I/O Pipes
winpty_conin_name = win_pty_dll.winpty_conin_name
winpty_conout_name = win_pty_dll.winpty_conout_name
winpty_conerr_name = win_pty_dll.winpty_conerr_name
# Agent RPC Calls
winpty_spawn_config_new = win_pty_dll.winpty_spawn_config_new
winpty_spawn_config_free = win_pty_dll.winpty_spawn_config_free
winpty_spawn = win_pty_dll.winpty_spawn
winpty_set_size = win_pty_dll.winpty_set_size
winpty_free = win_pty_dll.winpty_free
Kernel32 = ctypes.windll.Kernel32


def create_file(*args):
    return Kernel32.CreateFileW(*args)


def terminate_process(proc_h, exit_code):
    Kernel32.TerminateProcess(proc_h, exit_code)


def close_handle(h):
    if not h:
        Kernel32.CloseHandle(h)


def wait_for_single_object(h, milliseconds):
    res = Kernel32.WaitForSingleObject(h, milliseconds)
    if res == _winapi.WAIT_OBJECT_0:
        return
    elif res == _winapi.WAIT_TIMEOUT:
        raise TimeoutError("timeout")
    elif res == 0x00000080:
        raise Exception
    elif res == 0xFFFFFFFF:
        raise Exception("invalid process handle")


def env_dict2str(env_dict):
    if env_dict is None:
        return None
    temp = []
    for k, v in env_dict.items():
        temp.append("%s=%s" % (k, v))
    temp_new_env = "\0".join(temp)
    new_env = bytearray(temp_new_env, encoding='utf8')
    new_env.append(0)
    new_env.append(0)
    arr = (ctypes.c_uint16 * len(new_env))()
    for i in range(len(new_env)):
        arr[i] = new_env[i]
    return arr


class Process:

    def __init__(self, pty, proc_h, thread_h=None, stdin_h=None, stdout_h=None, stderr_h=None):
        self.pty = pty
        self.proc_h = proc_h
        self.thread_h = thread_h
        self.stdin_h = stdin_h
        self.stdout_h = stdout_h
        self.stderr_h = stderr_h
        self.closed = False
        self.out_str = None
        self._is_killed = False
        self.stdin = None if stdin_h is None else open(open_osfhandle(stdin_h, os.O_WRONLY), "wb")
        self.stdout = None if stdout_h is None else open(open_osfhandle(stdout_h, os.O_RDONLY), 'rb')
        self.stderr = None if stderr_h is None else open(open_osfhandle(stderr_h, os.O_RDONLY), 'rb')

    def wait(self, timeout=None):
        timeout = timeout or _winapi.INFINITE
        wait_for_single_object(self.proc_h, timeout * 1000)

    def readall(self, is_print=True, writer=None):
        try:
            self._reading(is_print, writer)
        except:
            pass
        self.stdin.close()

    def _reading(self, is_print=True, writer=None):
        stdout = self.stdout
        buffer = bytearray()
        while 1:
            data = stdout.read(1)
            if not data:
                break
            if data == b"\x1b":
                while 1:
                    data = stdout.read(1)
                    if data == b'\x07' or not data:
                        break
                continue
            if is_print:
                buffer.extend(data)
                if data[0] < 128:
                    t = buffer.decode("utf8")
                    print(t, end="")
                    buffer.clear()
            if writer:
                writer.write(data)
                writer.flush()

    def interactive(self):
        Thread(target=self.readall).start()
        file = self.stdin
        while 1:
            cmd_str = "%s\r\n" % input()
            if file.closed:
                break
            file.write(cmd_str.encode('utf8'))
            file.flush()
        self.close()

    def close(self):
        if self.closed:
            return
        self.getoutput()
        close_handle(self.stdin)
        close_handle(self.stdout)
        close_handle(self.stderr)
        close_handle(self.stderr)
        close_handle(self.thread_h)
        close_handle(self.proc_h)
        winpty_free(self.pty)
        self.closed = True

    def getoutput(self):
        if self.out_str is None:
            self.out_str = self.stdout.read()
        return self.out_str

    def kill(self):
        if self._is_killed:
            return
        terminate_process(self.proc_h, -1)
        self.close()
        self._is_killed = True

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


def create_process(cmdline: str, cwd=None, env_dict=None, winpty_flags=WINPTY_FLAG_PLAIN_OUTPUT):
    exec_file = None
    env = env_dict2str(env_dict)
    agent_cfg = winpty_config_new(winpty_flags, None)
    agent_cfg = ctypes.c_void_p(agent_cfg)
    assert agent_cfg is not None, "create agent config failed"
    pty = winpty_open(agent_cfg, None)
    pty = ctypes.c_void_p(pty)
    assert pty is not None, "open winpty failed"
    winpty_config_free(agent_cfg)
    stdin_handle = create_file(winpty_conin_name(pty), _winapi.GENERIC_WRITE, 0, None, _winapi.OPEN_EXISTING, 0, None)
    assert stdin_handle > 0, "the invalid stdin handle"
    stdout_handle = create_file(winpty_conout_name(pty), _winapi.GENERIC_READ, 0, None, _winapi.OPEN_EXISTING, 0, None)
    assert stdout_handle > 0, "invalid stdout handle"
    if winpty_flags & WINPTY_FLAG_CONERR:
        stderr_handle = create_file(winpty_conerr_name(pty), _winapi.GENERIC_READ, 0, None, _winapi.OPEN_EXISTING, 0,
                                    None)
        assert stdout_handle > 0, "invalid stderr handle"
    else:
        stderr_handle = None
    spawn_cfg = winpty_spawn_config_new(WINPTY_SPAWN_FLAG_AUTO_SHUTDOWN, exec_file, cmdline, cwd, env, None)
    spawn_cfg = ctypes.c_void_p(spawn_cfg)
    assert spawn_cfg is not None, "create spawn config failed"
    process = wintypes.HANDLE()
    thread = wintypes.HANDLE()
    spawn_success = winpty_spawn(pty, spawn_cfg, ctypes.pointer(process), ctypes.pointer(thread), None, None)
    assert spawn_success != 0 and process.value is not None, "create process failed"
    return Process(pty, process.value, thread.value, stdin_handle, stdout_handle, stderr_handle)


if __name__ == '__main__':
    cmd = "cmd.exe"
    # cmd = 'python'
    # cmd = 'powershell'
    # cmd = 'wmic'
    # cmd = 'ftp'
    # cmd = 'diskpart'
    # cmd = 'cmd /c "echo asd"'
    with create_process(cmd) as p:
        h = winpty_agent_process(p.pty)
        print(h)
        print(p.proc_h)
        p.interactive()
        try:
            p.wait(2)
        except:
            p.kill()
        print(p.getoutput().decode('utf8'))
