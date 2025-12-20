using WindowsInput;
using WindowsInput.Native;

namespace Office_Tools_Lite.Task_Helper;

public class KeyboardSimulator
{
    public InputSimulator Sim { get; private set; }
    public VirtualKeyCode Alt { get; private set; }
    public VirtualKeyCode A_Key { get; private set; }
    public VirtualKeyCode Enter { get; private set; }
    public VirtualKeyCode LeftCtrl { get; private set; }
    public VirtualKeyCode End_Key { get; private set; }
    public VirtualKeyCode Space { get; private set; }
    public VirtualKeyCode Backspace { get; private set; }
    public VirtualKeyCode F4_key { get; private set; }
    public VirtualKeyCode F5_key { get; private set; }
    public VirtualKeyCode F6_key { get; private set; }
    public VirtualKeyCode F7_key { get; private set; }
    public VirtualKeyCode F8_key { get; private set; }
    public VirtualKeyCode F9_key { get; private set; } 
    public VirtualKeyCode F10_key { get; private set; } 
    public VirtualKeyCode H_key { get; private set; }
    public VirtualKeyCode Down_key { get; private set; }
    public VirtualKeyCode Esc_key { get; private set; }
    public IKeyboardSimulator Input { get; private set; }

    // Constructor to initialize the properties
    public KeyboardSimulator()
    {
        Sim = new InputSimulator();
        Alt = VirtualKeyCode.LMENU;
        A_Key = VirtualKeyCode.VK_A;
        Enter = VirtualKeyCode.RETURN;
        LeftCtrl = VirtualKeyCode.LCONTROL;
        End_Key = VirtualKeyCode.END;
        Backspace = VirtualKeyCode.BACK;
        Space = VirtualKeyCode.SPACE;
        F4_key = VirtualKeyCode.F4;
        F5_key = VirtualKeyCode.F5;
        F6_key = VirtualKeyCode.F6;
        F7_key = VirtualKeyCode.F7;
        F8_key = VirtualKeyCode.F8;
        F9_key = VirtualKeyCode.F9; 
        F10_key = VirtualKeyCode.F10; 
        H_key = VirtualKeyCode.VK_H;
        Down_key = VirtualKeyCode.DOWN;
        Esc_key = VirtualKeyCode.ESCAPE;

        Input = Sim.Keyboard;
    }

    // Method to simulate pressing a key combination (example)

    //public void Alt_Down() { Input.KeyDown(Alt); }
    //public void Alt_Up() { Input.KeyUp(Alt); }
    //public void Ctrl_Down() { Input.KeyDown(LeftCtrl); }
    //public void Ctrl_Up() { Input.KeyUp(LeftCtrl); }
    //public void A_key() { Input.KeyPress(A_Key); }

    public void Alt_A() { Input.ModifiedKeyStroke(Alt, A_Key); }
    public void Ctrl_A() { Input.ModifiedKeyStroke(LeftCtrl, A_Key); }
    public void Ctrl_H() { Input.ModifiedKeyStroke(LeftCtrl, H_key); }
    public void Alt_F5() { Input.ModifiedKeyStroke(Alt, F5_key); }
    public void Alt_F6() { Input.ModifiedKeyStroke(Alt, F6_key); }

    public void Enter_key() { Input.KeyPress(Enter); }
    public void End_key() { Input.KeyPress(End_Key); }
    public void Ctrl_End() { Input.ModifiedKeyStroke(LeftCtrl, End_Key); }
    public void Spacebar() { Input.KeyPress(Space); }
    public void BackSpace_key() { Input.KeyPress(Backspace); }
    public void F4() { Input.KeyPress(F4_key); }
    public void F5() { Input.KeyPress(F5_key); }
    public void F6() { Input.KeyPress(F6_key); }
    public void F7() { Input.KeyPress(F7_key); }
    public void F8() { Input.KeyPress(F8_key); }
    public void F9() { Input.KeyPress(F9_key); } 
    public void F10() { Input.KeyPress(F10_key); }
    public void Down() { Input.KeyPress(Down_key); }
    public void Esc() { Input.KeyPress(Esc_key); }

}
