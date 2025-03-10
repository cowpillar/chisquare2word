import pyautogui
import time

def type_equation(equation):
    pyautogui.typewrite(equation)

def round_to_4_digits(value):
    return round(value, 4)

def main():
    o = []
    e = []
    oe = []
    oe_squared = []
    oe_squared_divided_by_e = []
    sum_of_oe_squared_divided_by_e = 0
    
    print("Welcome to 'Chi-Square to Word' program!")
    print("The purpose of this program is to efficiently compute and automate user input to Word.")
    print("In other words, it solves and displays chi-square problems automatically.")
    print(" ")

    # Observed Values
    print("Type the Observed Values.")
    print(" ")
    obs_input = input("Enter observed values separated by space: ")
    if obs_input.lower() != 'done':
        o = [float(val) for val in obs_input.split()]

    # Expected Values
    print(" ")
    print("Type the Expected Values.")
    print(" ")
    exp_input = input("Enter expected values separated by space: ")
    e = [float(val) for val in exp_input.split()]

    # Calculate differences
    oe = [obs - exp for obs, exp in zip(o, e)]
    oe_squared = [diff ** 2 for diff in oe]
    oe_squared_divided_by_e = [diff_squared / exp if exp != 0 else 0 for diff_squared, exp in zip(oe_squared, e)]
    sum_of_oe_squared_divided_by_e = round_to_4_digits(sum(oe_squared_divided_by_e))

    print(" ")
    print("Observed Values:", [int(val) if val.is_integer() else val for val in o])
    print("Expected Values:", [int(val) if val.is_integer() else val for val in e])
    print("O - E:", [int(d) if d.is_integer() else round_to_4_digits(d) for d in oe])
    print("(O - E)²:", [int(d) if d.is_integer() else round_to_4_digits(d) for d in oe_squared])
    print("(O - E)² / E:", [int(d) if d.is_integer() else round_to_4_digits(d) for d in oe_squared_divided_by_e])
    print("∑ (O - E)² / E:", round_to_4_digits(sum_of_oe_squared_divided_by_e))
    print(" ")

    # Formula
    equation = "x^2 = " + " + ".join([f"({int(oi) if oi.is_integer() else oi}-{int(ei) if ei.is_integer() else ei})^2/{int(ei) if ei.is_integer() else ei}" for oi, ei in zip(o, e)])
    # Equation for O - E
    equation2 = "x^2 = " + " + ".join([f"({int(d) if d.is_integer() else round_to_4_digits(d)})^2/{int(ei) if ei.is_integer() else ei}" for d, ei in zip(oe, e)])
    # Equation for (O - E)^2
    equation3 = "x^2 = " + " + ".join([f"{int(d) if d.is_integer() else round_to_4_digits(d)}/{int(ei) if ei.is_integer() else ei}" for d, ei in zip(oe_squared, e)])
    # Equation for (O - E)^2 / E
    equation4 = "x^2 = " + " + ".join([f"{int(d) if d.is_integer() else round_to_4_digits(d)}" for d in oe_squared_divided_by_e])
    # Equation for ∑ (O - E)^2 / E
    equation5 = f"x^2 = {round_to_4_digits(sum_of_oe_squared_divided_by_e)}"

    # Microsoft Word
    pyautogui.hotkey('winleft', 'r')
    time.sleep(1)
    pyautogui.typewrite('winword')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('alt', 'n')
    pyautogui.press('l')

    pyautogui.hotkey('alt', 'n')
    pyautogui.press('e')
    pyautogui.press('i')
    type_equation(equation)
    pyautogui.press('enter')

    pyautogui.hotkey('alt', 'n')
    pyautogui.press('e')
    pyautogui.press('i')
    type_equation(equation2)
    pyautogui.press('enter')

    pyautogui.hotkey('alt', 'n')
    pyautogui.press('e')
    pyautogui.press('i')
    type_equation(equation3)
    pyautogui.press('enter')

    pyautogui.hotkey('alt', 'n')
    pyautogui.press('e')
    pyautogui.press('i')
    type_equation(equation4)
    pyautogui.press('enter')

    pyautogui.hotkey('alt', 'n')
    pyautogui.press('e')
    pyautogui.press('i')
    type_equation(equation5)
    pyautogui.press('enter')

if __name__ == "__main__":
    main()
