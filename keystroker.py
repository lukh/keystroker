import win32com.client
import win32api, win32con
import rtmidi_python as rtmidi
import json
import getopt
import sys

CONFIG_FILE = "config.json"
SETTINGS_FILE = "settings.json"

# init the shell
shell = win32com.client.Dispatch("WScript.Shell")


def sendKey(key):
    shell.SendKeys(key, 0)


def wheel(val, shift=False, ctrl=False):
    if ctrl:
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    if shift:
        win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)


    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0,val,0)

    win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)



def loadJSON():
    config = {}
    try:
        with open(CONFIG_FILE) as f:
            config = json.load(f)
    except:
        print "No config.json file found. Consider using -a option"
        exit(0)

    return config




def learning(midi_in):
    print "entering learning mode. press a midi signal (NoteOn or CC) and then enter the keystroke you want"
    print "\nspecific keys\n"

    print '    SHIFT   +'
    print '    CTRL    ^'
    print '    ALT     %'
    print '    ENTER   ~'
    print '    BACKSPACE   \{BACKSPACE\}, \{BS\}, or \{BKSP\}'
    print '    BREAK   \{BREAK\}'
    print '    CAPS LOCK   \{CAPSLOCK\}'
    print '    DEL or DELETE   \{DELETE\} or \{DEL\}'
    print '    DOWN ARROW  \{DOWN\}'
    print '    END     \{END\}'
    print '    ENTER   \{ENTER\} or ~'
    print '    ESC     \{ESC\}'
    print '    HELP    \{HELP\}'
    print '    HOME    \{HOME\}'
    print '    INS or INSERT   \{INSERT\} or \{INS\}'
    print '    LEFT ARROW  \{LEFT\}'
    print '    NUM LOCK    \{NUMLOCK\}'
    print '    PAGE DOWN   \{PGDN\}'
    print '    PAGE UP     \{PGUP\}'
    print '    PRINT SCREEN    \{PRTSC\}'
    print '    RIGHT ARROW     \{RIGHT\}'
    print '    SCROLL LOCK     \{SCROLLLOCK\}'
    print '    TAB     \{TAB\}'
    print '    UP ARROW    \{UP\}\n'

    print "Mouse:"
    print "MOUSE;[UP,DOWN];(SHIFT,CTRL)\n"
    print "When done, hit Ctrl+C to kill and reload the software (without -a)"

    stop = False
    while not stop:
        #opening the file
        config = {}
        try:
            with open(CONFIG_FILE) as f:
                config = json.load(f)
        except:
            print "No config.json file found. Creating one\n\n"

        # catch the midi message
        print "Hit the MIDI knob/switch you want"
        message = []
        while True:
            message, delta_time = midi_in.get_message()
            if message:
                if message[0] != 128:
                    break
        if len(message) == 3:
            if message[0] == 144: # NoteOn 
                key = raw_input("enter keystroke:")
                config["{},{}".format(message[0], message[1])] = key
                print "    > learning: ", "{},{}".format(message[0], message[1]), " > ", key
            elif message[0] == 176:
                key = raw_input("enter keystroke:")
                config["{},{},{}".format(message[0], message[1], message[2])] = key
                print "    > learning: ", "{},{}".format(message[0], message[1], message[2]), " > ", key


        while message != None: # flush
            message, delta_time = midi_in.get_message()

        print "saving config..."
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f)



def handleMouseBinding(key_binding):
    splitted = key_binding.split(";")

    ctrl=False
    shift=False

    if len(splitted) >= 2:
        if len(splitted) == 3:
            magic_keys = splitted[2].split(',')
            shift = "SHIFT" in magic_keys
            ctrl = "CTRL" in magic_keys

        val = 120 if splitted[1] == "UP" else -120

        wheel(val, shift, ctrl)
        print val, shift, ctrl



def runtime(midi_in):
    print "running..."
    config = loadJSON()
    while True:
        message, delta_time = midi_in.get_message()
        if message:
            if len(message) == 3:
                if message[0] == 144:
                    k = config.get("{},{}".format(message[0], message[1]), "No binding")
                    
                    if k.find("MOUSE;") != -1:
                        handleMouseBinding(k)

                    elif k != "No binding":
                        print "key :", k
                        sendKey(k)
                elif message[0] == 176:
                    k = config.get("{},{},{}".format(message[0], message[1], message[2]), "No binding")
                    
                    if k.find("MOUSE;") != -1:
                        handleMouseBinding(k)

                    elif k != "No binding":
                        print "key :", k
                        sendKey(k)



def init_midi():
    midi_in = rtmidi.MidiIn()

    ports = midi_in.ports

    try:
        with open(SETTINGS_FILE) as f:
            settings = json.load(f)
            if "midi_port" not in settings:
                raise Exception()

            port_idx = settings["midi_port"]
    except:
        print "No settings.json file found. Creating one\n\n"
        print "Availables MIDI ports:"
        print "\n".join(["{}: {}".format(i,x) for i,x in enumerate(ports)])
        port_idx = int(raw_input("Which port do you want to use ?"))
        if port_idx >= len(ports):
            print "fuck off"
            exit(-42)

        #saving
        print "saving settings..."
        with open(SETTINGS_FILE, "w") as f:
            json.dump({"midi_port":port_idx}, f)


    print "Opening MIDI port", ports[port_idx]
    midi_in.open_port(port_idx)

    return midi_in

if __name__ == "__main__":
    learning_mode = False

    argv = sys.argv[1:]
    try:
        opts, args = getopt.getopt(argv,"am:",[])
    except getopt.GetoptError:
        print 'opts error !'
        print 'use -a option to add a new keystroke'
        sys.exit(2)

    for opt, arg in opts:
        if opt == '-h':
            print 'use -a option to add a new keystroke'
            sys.exit()
        elif opt in ("-a"):
            learning_mode = True



    # init midi 
    midi_in = init_midi()


    if learning_mode:
        learning(midi_in)


    runtime(midi_in)



