import React, { useState, useEffect, useRef } from "react";
import { PieChart, Pie, Cell, ResponsiveContainer } from "recharts";
import { getMemberDisplayName } from "./src/memberName";
import * as XLSX from "xlsx";
import {
  Lock,
  Unlock,
  CheckCircle2,
  XCircle,
  Clock,
  ArrowRight,
  ArrowLeft,
  ClipboardList,
  Trophy,
  RotateCcw,
  Terminal,
  BookOpen,
  CalendarDays,
  KeyRound,
  Code2,
  User,
  LogOut,
  Play,
  Trash2,
  ChevronDown,
  Loader2,
  Download,
} from "lucide-react";

const COLORS = {
  ink: "#1a1a1a",
  paper: "#f3eee1",
  paperDim: "#b8b0a0",
  inkSoft: "#2a2a2a",
  line: "#d4cbb8",
  green: "#3f8f63",
  clay: "#c1583f",
  amber: "#e8a338",
  teal: "#4a9b8f",
};

const CATEGORY_COLOR = {
  aptitude: "#3f8f63",
  verbal: "#e8a338",
  technical: "#4a9b8f",
  technical_c: "#c1583f",
  technical_java: "#8b5cf6",
  assignment_c: "#c1583f",
};

const FONT_IMPORT = `
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&family=Spectral:wght@400;600;700&display=swap');
`;

const LOGO_DATA_URI = "/EDU TECH.jpg";

const TECH_POOL = {
  easy: [
    {
      q: `What is the output of the following code?`,
      code: `x = 10
x += 5
print(x)`,
      opts: [`0`, `15`, `None`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `x = 10
x -= 3
print(x)`,
      opts: [`None`, `Error`, `7`, `0`],
      correct: 2,
      explanation: `The -= operator subtracts the right operand from the left operand and assigns the result to the left operand. x -= 3 is equivalent to x = x - 3. Since x = 10, x = 10 - 3 = 7. Therefore, print(x) outputs 7.`,
    },
    {
      q: `What is the output of the following code?`,
      code: `x = 3
x *= 4
print(x)`,
      opts: [`Error`, `None`, `12`, `0`],
      correct: 2,
      explanation: `The *= operator multiplies the left operand by the right operand and assigns the result to the left operand. x *= 4 is equivalent to x = x * 4. Since x = 3, x = 3 * 4 = 12. Therefore, print(x) outputs 12.`,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(max(3, 7, 2))`,
      opts: [`0`, `Error`, `None`, `7`],
      correct: 3,
      explanation: `The max() function returns the largest value among the given arguments. Among 3, 7, and 2, the largest value is 7. Therefore, print(max(3, 7, 2)) outputs 7.`,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(min(3, 7, 2))`,
      opts: [`2`, `0`, `Error`, `None`],
      correct: 0,
      explanation: `The min() function returns the smallest value among the given arguments. Among 3, 7, and 2, the smallest value is 2. Therefore, print(min(3, 7, 2)) outputs 2.`,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(sum([1, 2, 3, 4]))`,
      opts: [`Error`, `10`, `None`, `0`],
      correct: 1,
      explanation: `The sum() function calculates the sum of all elements in an iterable. sum([1, 2, 3, 4]) = 1 + 2 + 3 + 4 = 10. Therefore, print(sum([1, 2, 3, 4])) outputs 10.`,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('Hello' + ' ' + 'World')`,
      opts: [`None`, `Error`, `0`, `Hello World`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('ab' * 3)`,
      opts: [`Error`, `0`, `ababab`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(3 > 2 and 5 > 4)`,
      opts: [`None`, `0`, `True`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(3 > 2 or 5 < 4)`,
      opts: [`None`, `True`, `0`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(not True)`,
      opts: [`0`, `False`, `Error`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(None == False)`,
      opts: [`None`, `0`, `Error`, `False`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(1 if 5 > 3 else 0)`,
      opts: [`Error`, `0`, `None`, `1`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `for i in range(3):
    print(i)`,
      opts: [
        `Error`,
        `None`,
        `0`,
        `0
1
2`,
      ],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `i = 0
while i < 3:
    print(i)
    i += 1`,
      opts: [
        `0
1
2`,
        `None`,
        `Error`,
        `0`,
      ],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `for i in range(1, 6, 2):
    print(i, end=' ')`,
      opts: [`None`, `1 3 5`, `0`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(list(range(3)))`,
      opts: [`None`, `[0, 1, 2]`, `0`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(list(range(2, 8, 2)))`,
      opts: [`[2, 4, 6]`, `None`, `Error`, `0`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = 'python'
print(s[-1])`,
      opts: [`n`, `0`, `Error`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = 'python'
print(s[1:4])`,
      opts: [`None`, `0`, `yth`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = 'python'
print(s[:3])`,
      opts: [`0`, `None`, `pyt`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = 'python'
print(s[::-1])`,
      opts: [`Error`, `0`, `nohtyp`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = 'Python'
print(s.upper())`,
      opts: [`Error`, `PYTHON`, `None`, `0`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = 'PYTHON'
print(s.lower())`,
      opts: [`Error`, `0`, `None`, `python`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = '  hi  '
print(s.strip())`,
      opts: [`0`, `None`, `hi`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = 'a,b,c'
print(s.split(','))`,
      opts: [`None`, `['a', 'b', 'c']`, `0`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('-'.join(['a', 'b', 'c']))`,
      opts: [`0`, `a-b-c`, `None`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('py' in 'python')`,
      opts: [`0`, `None`, `Error`, `True`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('java' in 'python')`,
      opts: [`0`, `Error`, `False`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `a, b = 1, 2
a, b = b, a
print(a, b)`,
      opts: [`2 1`, `None`, `0`, `Error`],
      correct: 0,
    },
    {
      q: `Which keyword is used to define a function in Python?`,
      opts: [`func`, `def`, `function`, `lambda`],
      correct: 1,
    },
    {
      q: `What is the correct file extension for Python files?`,
      opts: [`.pt`, `.pyt`, `.py`, `.pyth`],
      correct: 2,
    },
    {
      q: `Which of these is a mutable data type in Python?`,
      opts: [`List`, `Tuple`, `String`, `Int`],
      correct: 0,
    },
    {
      q: `What does the len() function do?`,
      opts: [
        `Returns the type of an object`,
        `Returns the length of an object`,
        `Returns the memory size`,
        `Deletes an object`,
      ],
      correct: 1,
    },
    {
      q: `Which operator is used for exponentiation in Python?`,
      opts: [`^`, `**`, `//`, `%%`],
      correct: 1,
    },
    {
      q: `Which list method adds an item to the end of a list?`,
      opts: [`insert()`, `add()`, `append()`, `extend()`],
      correct: 2,
    },
    {
      q: `Which construct is used to handle exceptions in Python?`,
      opts: [`if-else`, `try-except`, `switch-case`, `catch-throw`],
      correct: 1,
    },
    {
      q: `Which data structure does Python use to store key-value pairs?`,
      opts: [`List`, `Set`, `Dictionary`, `Tuple`],
      correct: 2,
    },
    {
      q: `Which function converts a string to an integer in Python?`,
      opts: [`str()`, `int()`, `float()`, `chr()`],
      correct: 1,
    },
    {
      q: `Which of these is NOT a built-in Python data type?`,
      opts: [`int`, `float`, `character`, `list`],
      correct: 2,
    },
    {
      q: `What does the 'pass' statement do in Python?`,
      opts: [
        `Ends a loop`,
        `Does nothing; acts as a placeholder`,
        `Skips to the next iteration`,
        `Raises an exception`,
      ],
      correct: 1,
    },
    {
      q: `Which keyword is used to create a class in Python?`,
      opts: [`object`, `class`, `struct`, `define`],
      correct: 1,
    },
    {
      q: `What is the key difference between a list and a tuple in Python?`,
      opts: [
        `Lists are mutable, tuples are immutable`,
        `Tuples are mutable, lists are immutable`,
        `Lists can't hold strings`,
        `Tuples can't hold numbers`,
      ],
      correct: 0,
    },
    {
      q: `Which keyword is used to import a module in Python?`,
      opts: [`include`, `import`, `using`, `require`],
      correct: 1,
    },
    {
      q: `Which of these is Python's standard package manager?`,
      opts: [`npm`, `pip`, `gem`, `composer`],
      correct: 1,
    },
  ],
  medium: [
    {
      q: `What is the output of the following code?`,
      code: `nums = [1, 2, 3, 4, 5]
print([n*n for n in nums if n % 2 == 0])`,
      opts: [`[4, 16]`, `TypeError`, `None`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `nums = [1, 2, 3]
print([n + 1 for n in nums])`,
      opts: [`None`, `[2, 3, 4]`, `Error`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `d = {'a': 1, 'b': 2}
print(d.get('c', 0))`,
      opts: [`None`, `TypeError`, `Error`, `0`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `d = {'a': 1, 'b': 2}
print(list(d.keys()))`,
      opts: [`Error`, `TypeError`, `None`, `['a', 'b']`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `d = {'a': 1, 'b': 2}
print(list(d.values()))`,
      opts: [`TypeError`, `Error`, `[1, 2]`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `d = {'a': 1}
d['b'] = 2
print(d)`,
      opts: [`None`, `Error`, `TypeError`, `{'a': 1, 'b': 2}`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1, 2, 3]
b = a
b.append(4)
print(a)`,
      opts: [`TypeError`, `Error`, `[1, 2, 3, 4]`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1, 2, 3]
b = a.copy()
b.append(4)
print(a)`,
      opts: [`Error`, `[1, 2, 3]`, `TypeError`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [3, 1, 2]
a.sort()
print(a)`,
      opts: [`Error`, `[1, 2, 3]`, `TypeError`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [3, 1, 2]
print(sorted(a, reverse=True))`,
      opts: [`None`, `[3, 2, 1]`, `TypeError`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1, 2, 3]
a.remove(2)
print(a)`,
      opts: [`Error`, `TypeError`, `None`, `[1, 3]`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1, 2, 3]
a.pop()
print(a)`,
      opts: [`TypeError`, `[1, 2]`, `Error`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1, 2, 3]
a.insert(1, 99)
print(a)`,
      opts: [`TypeError`, `Error`, `[1, 99, 2, 3]`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1, [2, 3], 4]
print(a[1][0])`,
      opts: [`None`, `2`, `Error`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = (1, 2, 3)
print(a[0])`,
      opts: [`TypeError`, `Error`, `None`, `1`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `t = (1, 2)
a, b = t
print(a + b)`,
      opts: [`Error`, `TypeError`, `3`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `s = {1, 2, 3}
s.add(2)
print(len(s))`,
      opts: [`TypeError`, `Error`, `3`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `s1 = {1, 2, 3}
s2 = {2, 3, 4}
print(s1 & s2)`,
      opts: [`{2, 3}`, `TypeError`, `None`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `s1 = {1, 2, 3}
s2 = {2, 3, 4}
print(s1 | s2)`,
      opts: [`Error`, `{1, 2, 3, 4}`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `s1 = {1, 2, 3}
s2 = {2, 3}
print(s1 - s2)`,
      opts: [`Error`, `None`, `TypeError`, `{1}`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `def add(a, b):
    return a + b
print(add(3, 4))`,
      opts: [`None`, `TypeError`, `Error`, `7`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `def greet(name='World'):
    return f'Hello, {name}!'
print(greet())`,
      opts: [`Error`, `None`, `Hello, World!`, `TypeError`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `def f(*args):
    return sum(args)
print(f(1, 2, 3))`,
      opts: [`TypeError`, `6`, `Error`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `def f(**kwargs):
    return len(kwargs)
print(f(a=1, b=2, c=3))`,
      opts: [`Error`, `TypeError`, `3`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `def outer():
    x = 1
    def inner():
        return x + 1
    return inner()
print(outer())`,
      opts: [`TypeError`, `None`, `Error`, `2`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `square = lambda x: x * x
print(square(6))`,
      opts: [`TypeError`, `None`, `36`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `nums = [1, 2, 3, 4]
print(list(map(lambda x: x * 2, nums)))`,
      opts: [`None`, `[2, 4, 6, 8]`, `TypeError`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `nums = [1, 2, 3, 4, 5, 6]
print(list(filter(lambda x: x % 2 == 0, nums)))`,
      opts: [`[2, 4, 6]`, `TypeError`, `None`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `from functools import reduce
nums = [1, 2, 3, 4]
print(reduce(lambda a, b: a + b, nums))`,
      opts: [`Error`, `10`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(sorted(['banana', 'apple', 'cherry']))`,
      opts: [`Error`, `TypeError`, `['apple', 'banana', 'cherry']`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `words = ['a', 'bb', 'ccc']
print(sorted(words, key=len))`,
      opts: [`None`, `Error`, `['a', 'bb', 'ccc']`, `TypeError`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `def fact(n):
    if n <= 1:
        return 1
    return n * fact(n - 1)
print(fact(5))`,
      opts: [`None`, `TypeError`, `__ERROR__:NameError`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `def fib(n):
    if n <= 1:
        return n
    return fib(n-1) + fib(n-2)
print(fib(6))`,
      opts: [`TypeError`, `__ERROR__:NameError`, `Error`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `try:
    print(10 / 0)
except ZeroDivisionError:
    print('Cannot divide by zero')`,
      opts: [`TypeError`, `None`, `Cannot divide by zero`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `try:
    x = int('abc')
except ValueError:
    print('Invalid number')`,
      opts: [`Invalid number`, `None`, `Error`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `try:
    print(1)
finally:
    print(2)`,
      opts: [
        `TypeError`,
        `None`,
        `Error`,
        `1
2`,
      ],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `for i in range(5):
    if i == 3:
        break
    print(i, end='')`,
      opts: [`012`, `Error`, `None`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `for i in range(5):
    if i == 2:
        continue
    print(i, end='')`,
      opts: [`None`, `Error`, `TypeError`, `0134`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `for i in range(3):
    for j in range(2):
        print(i, j, end=' ')`,
      opts: [`0 0 0 1 1 0 1 1 2 0 2 1`, `TypeError`, `None`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print([x for x in range(10) if x % 3 == 0])`,
      opts: [`None`, `Error`, `TypeError`, `[0, 3, 6, 9]`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print({x: x*x for x in range(4)})`,
      opts: [`{0: 0, 1: 1, 2: 4, 3: 9}`, `Error`, `TypeError`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(any([False, False, True]))`,
      opts: [`Error`, `True`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(all([True, True, False]))`,
      opts: [`Error`, `False`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1, 2, 3]
b = [4, 5, 6]
print(list(zip(a, b)))`,
      opts: [`TypeError`, `Error`, `[(1, 4), (2, 5), (3, 6)]`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `for i, v in enumerate(['a', 'b', 'c']):
    print(i, v, end=' ')`,
      opts: [`Error`, `TypeError`, `None`, `0 a 1 b 2 c`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('{} and {}'.format('cats', 'dogs'))`,
      opts: [`TypeError`, `cats and dogs`, `Error`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `name = 'World'
print(f'Hello, {name}!')`,
      opts: [`Hello, World!`, `Error`, `TypeError`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('%d apples' % 5)`,
      opts: [`TypeError`, `Error`, `None`, `5 apples`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Dog:
    def __init__(self, name):
        self.name = name
d = Dog('Rex')
print(d.name)`,
      opts: [`Rex`, `Error`, `TypeError`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Counter:
    count = 0
    def __init__(self):
        Counter.count += 1
Counter(); Counter()
print(Counter.count)`,
      opts: [`__ERROR__:NameError`, `TypeError`, `Error`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Animal:
    def speak(self):
        return 'generic sound'
class Cat(Animal):
    def speak(self):
        return 'Meow'
print(Cat().speak())`,
      opts: [`None`, `TypeError`, `Meow`, `Error`],
      correct: 2,
    },
    {
      q: `What is a lambda function in Python?`,
      opts: [
        `A named function stored in a module`,
        `An anonymous function defined with the lambda keyword`,
        `A function that only returns None`,
        `A built-in decorator`,
      ],
      correct: 1,
    },
    {
      q: `What does the 'self' keyword represent in a class method?`,
      opts: [
        `The class itself`,
        `The instance of the class`,
        `A global variable`,
        `The parent class`,
      ],
      correct: 1,
    },
    {
      q: `Which module is used for regular expressions in Python?`,
      opts: [`regex`, `re`, `pyregex`, `restring`],
      correct: 1,
    },
    {
      q: `Which of these is a valid list comprehension in Python?`,
      opts: [
        `[x for x in range(10)]`,
        `for x in range(10) [x]`,
        `list(x for 10)`,
        `{x in range(10)}`,
      ],
      correct: 0,
    },
    {
      q: `What does the __init__ method do in a Python class?`,
      opts: [
        `Deletes an object`,
        `Initializes a new object's attributes`,
        `Creates a static method`,
        `Imports the class`,
      ],
      correct: 1,
    },
    {
      q: `What does PEP 8 refer to in Python?`,
      opts: [
        `A testing framework`,
        `Python's official style guide`,
        `A package manager`,
        `A data structure`,
      ],
      correct: 1,
    },
    {
      q: `Which of these best describes a Python virtual environment?`,
      opts: [
        `A cloud-only Python runtime`,
        `An isolated space with its own installed packages`,
        `A type of Python class`,
        `A GUI framework`,
      ],
      correct: 1,
    },
    {
      q: `What does the \`global\` keyword do inside a function?`,
      opts: [
        `Creates a new local variable`,
        `Allows the function to modify a variable from the module scope`,
        `Deletes a variable`,
        `Imports a global module`,
      ],
      correct: 1,
    },
  ],
  advanced: [
    {
      q: `What is the output of the following code?`,
      code: `def gen():
    yield 1
    yield 2
    yield 3
g = gen()
print(next(g))`,
      opts: [`Error`, `1`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `def gen():
    for i in range(3):
        yield i * i
print(list(gen()))`,
      opts: [`TypeError`, `[0, 1, 4]`, `Error`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `def counter():
    count = 0
    def inc():
        nonlocal count
        count += 1
        return count
    return inc
c = counter()
c(); c()
print(c())`,
      opts: [`TypeError`, `3`, `Error`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `def decorator(func):
    def wrapper():
        return func().upper()
    return wrapper
@decorator
def greet():
    return 'hello'
print(greet())`,
      opts: [`HELLO`, `Error`, `None`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `def repeat(n):
    def decorator(func):
        def wrapper(*args):
            for _ in range(n):
                func(*args)
        return wrapper
    return decorator
@repeat(2)
def say(msg):
    print(msg, end=' ')
say('hi')`,
      opts: [`hi hi`, `None`, `Error`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Base:
    def __init__(self):
        self.value = 10
class Derived(Base):
    def __init__(self):
        super().__init__()
        self.value += 5
print(Derived().value)`,
      opts: [`15`, `TypeError`, `None`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `class A:
    def __str__(self):
        return 'A instance'
print(str(A()))`,
      opts: [`TypeError`, `A instance`, `None`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Vector:
    def __init__(self, x, y):
        self.x, self.y = x, y
    def __add__(self, other):
        return Vector(self.x + other.x, self.y + other.y)
    def __repr__(self):
        return f'({self.x}, {self.y})'
print(Vector(1,2) + Vector(3,4))`,
      opts: [`__ERROR__:NameError`, `Error`, `TypeError`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Meta:
    @staticmethod
    def double(x):
        return x * 2
print(Meta.double(21))`,
      opts: [`None`, `Error`, `TypeError`, `42`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Circle:
    def __init__(self, r):
        self._r = r
    @property
    def area(self):
        return 3.14 * self._r ** 2
print(Circle(2).area)`,
      opts: [`TypeError`, `None`, `12.56`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print([i for i in range(20) if all(i % j != 0 for j in range(2, i))][1:5])`,
      opts: [`Error`, `None`, `TypeError`, `[1, 2, 3, 5]`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `matrix = [[1,2],[3,4]]
print([row[i] for row in matrix for i in range(len(row))])`,
      opts: [`Error`, `[1, 2, 3, 4]`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(list(zip(*[[1,2,3],[4,5,6]])))`,
      opts: [`None`, `[(1, 4), (2, 5), (3, 6)]`, `Error`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `from collections import Counter
print(Counter('mississippi').most_common(1))`,
      opts: [`[('i', 4)]`, `None`, `Error`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `from collections import defaultdict
d = defaultdict(int)
d['a'] += 1
d['a'] += 1
print(d['a'])`,
      opts: [`TypeError`, `None`, `2`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(sorted([(1,'b'), (2,'a')], key=lambda x: x[1]))`,
      opts: [`None`, `TypeError`, `Error`, `[(2, 'a'), (1, 'b')]`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(id([1,2]) == id([1,2]))`,
      opts: [`True`, `TypeError`, `None`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1, 2, 3]
b = a[:]
print(a is b)`,
      opts: [`Error`, `TypeError`, `None`, `False`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(2 is 2)`,
      opts: [`None`, `TypeError`, `Error`, `True`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `x = 5
def f():
    x = 10
    return x
print(f(), x)`,
      opts: [`10 5`, `Error`, `None`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `def f(a, b=[]):
    b.append(a)
    return b
f(1)
print(f(2))`,
      opts: [`[1, 2]`, `Error`, `None`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print((lambda *a: sum(a))(1,2,3,4))`,
      opts: [`None`, `Error`, `10`, `TypeError`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(*[1,2,3])`,
      opts: [`None`, `1 2 3`, `Error`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `def f(a, *, b):
    return a + b
print(f(1, b=2))`,
      opts: [`3`, `Error`, `TypeError`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(list(enumerate(['x','y'], start=1)))`,
      opts: [`[(1, 'x'), (2, 'y')]`, `None`, `TypeError`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `import itertools
print(list(itertools.permutations([1,2], 2)))`,
      opts: [`TypeError`, `Error`, `None`, `[(1, 2), (2, 1)]`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `import itertools
print(list(itertools.combinations([1,2,3], 2)))`,
      opts: [`[(1, 2), (1, 3), (2, 3)]`, `None`, `TypeError`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(bin(10))`,
      opts: [`Error`, `0b1010`, `TypeError`, `None`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(hex(255))`,
      opts: [`TypeError`, `None`, `0xff`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(oct(8))`,
      opts: [`Error`, `0o10`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(int('1010', 2))`,
      opts: [`None`, `TypeError`, `10`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(int('ff', 16))`,
      opts: [`255`, `None`, `Error`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(5 & 3)`,
      opts: [`Error`, `None`, `1`, `TypeError`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(5 | 2)`,
      opts: [`None`, `7`, `Error`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(5 ^ 1)`,
      opts: [`None`, `4`, `Error`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(~5)`,
      opts: [`-6`, `TypeError`, `Error`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(1 << 3)`,
      opts: [`Error`, `8`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(16 >> 2)`,
      opts: [`None`, `Error`, `4`, `TypeError`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('a' < 'b')`,
      opts: [`None`, `True`, `Error`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(ord('A'))`,
      opts: [`65`, `None`, `Error`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(chr(97))`,
      opts: [`None`, `a`, `TypeError`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(divmod(17, 5))`,
      opts: [`None`, `TypeError`, `Error`, `(3, 2)`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `class A:
    x = 1
class B(A):
    pass
B.x = 2
print(A.x, B.x)`,
      opts: [`TypeError`, `Error`, `None`, `1 2`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `try:
    raise ValueError('bad')
except Exception as e:
    print(type(e).__name__, e)`,
      opts: [`None`, `ValueError bad`, `TypeError`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `def f():
    try:
        return 1
    finally:
        print('cleanup')
print(f())`,
      opts: [
        `TypeError`,
        `None`,
        `Error`,
        `cleanup
1`,
      ],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print([1,2,3] == [1,2,3])`,
      opts: [`TypeError`, `Error`, `True`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print([1,2,3] is [1,2,3])`,
      opts: [`False`, `TypeError`, `Error`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(tuple([1,2,3]))`,
      opts: [`(1, 2, 3)`, `TypeError`, `Error`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(set([1,1,2,2,3]))`,
      opts: [`None`, `TypeError`, `Error`, `{1, 2, 3}`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(dict(zip(['a','b'], [1,2])))`,
      opts: [`None`, `TypeError`, `Error`, `{'a': 1, 'b': 2}`],
      correct: 3,
    },
    {
      q: `What is the purpose of the 'with' statement in Python?`,
      opts: [
        `Loop iteration`,
        `Context management, like automatic file closing`,
        `Exception raising`,
        `Class inheritance`,
      ],
      correct: 1,
    },
    {
      q: `Which keyword marks a function as a generator?`,
      opts: [`return`, `yield`, `async`, `gen`],
      correct: 1,
    },
    {
      q: `What is a Python decorator primarily used for?`,
      opts: [
        `Styling console output`,
        `Wrapping a function to extend its behavior`,
        `Declaring constants`,
        `Importing modules`,
      ],
      correct: 1,
    },
    {
      q: `What is the time complexity of a dictionary lookup in Python on average?`,
      opts: [`O(n)`, `O(log n)`, `O(1)`, `O(n^2)`],
      correct: 2,
    },
    {
      q: `Which built-in function pauses execution and returns control to the caller repeatedly?`,
      opts: [`return`, `yield`, `break`, `pass`],
      correct: 1,
    },
  ],
  pro: [
    {
      q: `What is the output of the following code?`,
      code: `def memo(f):
    cache = {}
    def wrapper(n):
        if n not in cache:
            cache[n] = f(n)
        return cache[n]
    return wrapper
@memo
def square(n):
    return n * n
print(square(5), square(5))`,
      opts: [`Error`, `TypeError`, `None`, `25 25`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Singleton:
    _instance = None
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
print(Singleton() is Singleton())`,
      opts: [`Error`, `None`, `TypeError`, `True`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `class Meta(type):
    def __call__(cls, *args):
        print('Creating instance')
        return super().__call__(*args)
class A(metaclass=Meta):
    pass
A()`,
      opts: [`Error`, `TypeError`, `None`, `Creating instance`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `def context():
    print('enter')
    yield 42
    print('exit')
from contextlib import contextmanager
cm = contextmanager(context)
with cm() as val:
    print(val)`,
      opts: [
        `None`,
        `TypeError`,
        `enter
42
exit`,
        `Error`,
      ],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(3.1 + 3.2 == 6.3)`,
      opts: [`Error`, `TypeError`, `None`, `False`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(0.1 + 0.2)`,
      opts: [`0.30000000000000004`, `Error`, `None`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `import sys
print(sys.maxsize > 0)`,
      opts: [`Error`, `TypeError`, `True`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [1,2,3]
del a[1]
print(a)`,
      opts: [`None`, `[1, 3]`, `TypeError`, `Error`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `a = [[0]*3]*3
a[0][0] = 1
print(a)`,
      opts: [`TypeError`, `None`, `Error`, `[[1, 0, 0], [1, 0, 0], [1, 0, 0]]`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print([x*2 for x in range(5)][::-1])`,
      opts: [`Error`, `[8, 6, 4, 2, 0]`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(''.join(reversed('hello')))`,
      opts: [`Error`, `TypeError`, `None`, `olleh`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('level' == 'level'[::-1])`,
      opts: [`TypeError`, `Error`, `None`, `True`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `def is_palindrome(s):
    return s == s[::-1]
print(is_palindrome('racecar'))`,
      opts: [`Error`, `TypeError`, `None`, `True`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `def flatten(lst):
    result = []
    for item in lst:
        if isinstance(item, list):
            result.extend(flatten(item))
        else:
            result.append(item)
    return result
print(flatten([1, [2, [3, 4], 5]]))`,
      opts: [`__ERROR__:NameError`, `Error`, `TypeError`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `def merge(d1, d2):
    return {**d1, **d2}
print(merge({'a':1}, {'a':2, 'b':3}))`,
      opts: [`TypeError`, `Error`, `{'a': 2, 'b': 3}`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(list(map(str, [1,2,3])))`,
      opts: [`TypeError`, `None`, `Error`, `['1', '2', '3']`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(sum(range(1, 101)))`,
      opts: [`TypeError`, `None`, `Error`, `5050`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(len(set([1,2,2,3,3,3])))`,
      opts: [`Error`, `3`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print([i for i in range(2, 30) if all(i % j for j in range(2, int(i**0.5)+1))][:5])`,
      opts: [`None`, `Error`, `[2, 3, 5, 7, 11]`, `TypeError`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `def quicksort(a):
    if len(a) <= 1:
        return a
    p = a[0]
    less = [x for x in a[1:] if x < p]
    more = [x for x in a[1:] if x >= p]
    return quicksort(less) + [p] + quicksort(more)
print(quicksort([3,6,1,8,2]))`,
      opts: [`__ERROR__:NameError`, `None`, `Error`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `def binary_search(a, target):
    lo, hi = 0, len(a)-1
    while lo <= hi:
        mid = (lo+hi)//2
        if a[mid] == target:
            return mid
        elif a[mid] < target:
            lo = mid+1
        else:
            hi = mid-1
    return -1
print(binary_search([1,2,3,4,5], 3))`,
      opts: [`2`, `1`, `3`, `-1`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `def fib_gen():
    a, b = 0, 1
    while True:
        yield a
        a, b = b, a + b
g = fib_gen()
print([next(g) for _ in range(7)])`,
      opts: [`[0, 1, 1, 2, 3, 5, 8]`, `None`, `TypeError`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(max(range(10), key=lambda x: -x))`,
      opts: [`TypeError`, `Error`, `0`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(min([3,1,2], default=0))`,
      opts: [`TypeError`, `Error`, `1`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(str.isdigit('123'))`,
      opts: [`True`, `None`, `TypeError`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('Hello World'.replace('World', 'Python'))`,
      opts: [`Hello Python`, `TypeError`, `Error`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('a-b-c'.split('-', 1))`,
      opts: [`Error`, `TypeError`, `None`, `['a', 'b-c']`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('   '.isspace())`,
      opts: [`None`, `Error`, `TypeError`, `True`],
      correct: 3,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('Hello'.startswith('He'))`,
      opts: [`Error`, `True`, `None`, `TypeError`],
      correct: 1,
    },
    {
      q: `What is the output of the following code?`,
      code: `print('Hello'.endswith('lo'))`,
      opts: [`True`, `Error`, `None`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(','.join(str(x) for x in range(5)))`,
      opts: [`Error`, `None`, `0,1,2,3,4`, `TypeError`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(list(filter(None, [0, 1, '', 'a', None, 2])))`,
      opts: [`TypeError`, `Error`, `[1, 'a', 2]`, `None`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `def f(x, y=10, *args, z=5, **kwargs):
    return x + y + z
print(f(1))`,
      opts: [`16`, `None`, `TypeError`, `Error`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `class A:
    __slots__ = ['x']
a = A()
a.x = 5
print(a.x)`,
      opts: [`5`, `Error`, `None`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(callable(print))`,
      opts: [`True`, `Error`, `None`, `TypeError`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(hasattr(str, 'upper'))`,
      opts: [`True`, `Error`, `TypeError`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(isinstance(5, (int, float)))`,
      opts: [`True`, `TypeError`, `Error`, `None`],
      correct: 0,
    },
    {
      q: `What is the output of the following code?`,
      code: `print(issubclass(bool, int))`,
      opts: [`TypeError`, `None`, `True`, `Error`],
      correct: 2,
    },
    {
      q: `What is the output of the following code?`,
      code: `async def foo():
    return 42
import asyncio
print(asyncio.run(foo()))`,
      opts: [`Error`, `TypeError`, `42`, `None`],
      correct: 2,
    },
    {
      q: `What does GIL stand for in CPython?`,
      opts: [
        `Global Instance Loader`,
        `Global Interpreter Lock`,
        `Generic Interface Layer`,
        `Garbage Iteration Logic`,
      ],
      correct: 1,
    },
    {
      q: `What does the walrus operator ':=' do in Python?`,
      opts: [
        `Assigns and returns a value in the same expression`,
        `Compares two values`,
        `Declares a constant`,
        `Creates a new class`,
      ],
      correct: 0,
    },
    {
      q: `Which module provides high-level threading and multiprocessing support?`,
      opts: [`os`, `threading / multiprocessing`, `sys`, `itertools`],
      correct: 1,
    },
    {
      q: `What is duck typing in Python?`,
      opts: [
        `A strict type-checking system`,
        `An object's suitability is determined by its methods, not its class`,
        `A way to declare constants`,
        `A type of loop`,
      ],
      correct: 1,
    },
    {
      q: `Which of these correctly describes Python's \`__str__\` vs \`__repr__\`?`,
      opts: [
        `They are identical in every way`,
        `__str__ is for readability, __repr__ is for unambiguous debugging output`,
        `__repr__ is only used in loops`,
        `__str__ can't be overridden`,
      ],
      correct: 1,
    },
    {
      q: `Which of these is true about Python's garbage collection?`,
      opts: [
        `Python never frees memory`,
        `Python uses reference counting plus a cyclic garbage collector`,
        `Memory must be freed manually`,
        `Python only runs on Windows`,
      ],
      correct: 1,
    },
  ],
};

const TECH_LEVEL_ORDER = ["easy", "medium", "hard"];

function buildTechnicalItems(seed, roundsPerLevel = 20) {
  const rng = mulberry32(hashStr(String(seed)));
  const shuffledLevels = {};
  for (const level of TECH_LEVEL_ORDER) {
    const arr = [...TECH_LEVEL_POOL[level]];
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(rng() * (i + 1));
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    shuffledLevels[level] = arr;
  }
  const items = [];
  for (let r = 0; r < roundsPerLevel; r++) {
    for (const level of TECH_LEVEL_ORDER) {
      const picked = shuffledLevels[level][r % shuffledLevels[level].length];
      items.push({ ...picked, tier: level });
    }
  }
  return items;
}

const TECH_LEVEL_POOL = {};

const TECH_QUESTIONS = Object.values(TECH_POOL).flatMap((arr) => arr);
const techLevelRng = mulberry32(hashStr("technical-levels"));
const shuffledQuestions = [...TECH_QUESTIONS];
for (let i = shuffledQuestions.length - 1; i > 0; i--) {
  const j = Math.floor(techLevelRng() * (i + 1));
  [shuffledQuestions[i], shuffledQuestions[j]] = [
    shuffledQuestions[j],
    shuffledQuestions[i],
  ];
}
const baseSize = Math.floor(shuffledQuestions.length / TECH_LEVEL_ORDER.length);
const extra = shuffledQuestions.length % TECH_LEVEL_ORDER.length;
TECH_LEVEL_ORDER.forEach((level, index) => {
  const size = baseSize + (index < extra ? 1 : 0);
  TECH_LEVEL_POOL[level] = shuffledQuestions.splice(0, size);
});

const TECH_POOL_FLAT = TECH_LEVEL_ORDER.flatMap(
  (level) => TECH_LEVEL_POOL[level],
);

// --- Technical (C) question bank: 60 questions (20 Easy, 20 Medium, 20 Hard) --
const C_POOL = {
  easy: [
    // 1
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { printf("%d", 2 + 3); return 0; }`,
      opts: [`2`, `3`, `5`, `Compiler error`],
      correct: 2,
    },
    // 2
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { printf("%d", 5 / 2); return 0; }`,
      opts: [`2`, `2.5`, `3`, `Compiler error`],
      correct: 0,
    },
    // 3
    {
      q: `Which header file is required to use printf() in C?`,
      opts: [`<stdlib.h>`, `<stdio.h>`, `<string.h>`, `<math.h>`],
      correct: 1,
    },
    // 4
    {
      q: `What is the correct file extension for a C source file?`,
      opts: [`.cpp`, `.c`, `.py`, `.cs`],
      correct: 1,
    },
    // 5
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int a = 10; a += 5; printf("%d", a); return 0; }`,
      opts: [`5`, `10`, `15`, `Compiler error`],
      correct: 2,
    },
    // 6
    {
      q: `Which keyword is used to define a constant in C?`,
      opts: [`var`, `constant`, `const`, `final`],
      correct: 2,
    },
    // 7
    {
      q: `What does the %d format specifier print in C?`,
      opts: [`A float`, `A character`, `A string`, `An integer`],
      correct: 3,
    },
    // 8
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { printf("%d", 10 % 3); return 0; }`,
      opts: [`0`, `1`, `3`, `Compiler error`],
      correct: 1,
    },
    // 9
    {
      q: `Which of the following correctly declares an integer variable in C?`,
      opts: [`integer x;`, `int x;`, `Int x;`, `var x;`],
      correct: 1,
    },
    // 10
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int x = 5; printf("%d", x * 2); return 0; }`,
      opts: [`5`, `7`, `10`, `Compiler error`],
      correct: 2,
    },
    // 11
    {
      q: `Which symbol is used to end a statement in C?`,
      opts: [`.`, `:`, `;`, `}`],
      correct: 2,
    },
    // 12
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { char c = 'A'; printf("%c", c); return 0; }`,
      opts: [`65`, `A`, `a`, `Compiler error`],
      correct: 1,
    },
    // 13
    {
      q: `What does the return type void mean for a function in C?`,
      opts: [
        `The function returns 0`,
        `The function returns nothing`,
        `The function returns a pointer`,
        `The function never ends`,
      ],
      correct: 1,
    },
    // 14
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int a = 4, b = 3; printf("%d", a - b); return 0; }`,
      opts: [`7`, `1`, `12`, `Compiler error`],
      correct: 1,
    },
    // 15
    {
      q: `Which of the following is NOT a valid C data type?`,
      opts: [`int`, `float`, `bool`, `string`],
      correct: 3,
    },
    // 16
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { printf("%f", 1.5 + 2.5); return 0; }`,
      opts: [`3`, `4.0`, `4.000000`, `Compiler error`],
      correct: 2,
    },
    // 17
    {
      q: `How do you write a single-line comment in C?`,
      opts: [`/* comment */`, `// comment`, `# comment`, `-- comment`],
      correct: 1,
    },
    // 18
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int x = 7; if (x > 5) printf("yes"); return 0; }`,
      opts: [`no`, `yes`, `7`, `Compiler error`],
      correct: 1,
    },
    // 19
    {
      q: `Which operator is used for assignment in C?`,
      opts: [`==`, `=>`, `=`, `:=`],
      correct: 2,
    },
    // 20
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int i; for(i = 0; i < 3; i++) printf("%d ", i); return 0; }`,
      opts: [`1 2 3`, `0 1 2`, `0 1 2 3`, `Compiler error`],
      correct: 1,
    },
  ],

  medium: [
    // 1
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int a = 3; int b = a++; printf("%d %d", a, b); return 0; }`,
      opts: [`4 3`, `3 4`, `3 3`, `4 4`],
      correct: 0,
    },
    // 2
    {
      q: `Which function is used to dynamically allocate memory in C?`,
      opts: [`alloc()`, `malloc()`, `new()`, `create()`],
      correct: 1,
    },
    // 3
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int a = 5; printf("%d", ++a); return 0; }`,
      opts: [`5`, `6`, `7`, `Compiler error`],
      correct: 1,
    },
    // 4
    {
      q: `Which of the following correctly declares a pointer to an int in C?`,
      opts: [`int p;`, `int *p;`, `p int*;`, `pointer<int> p;`],
      correct: 1,
    },
    // 5
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int arr[3] = {10, 20, 30}; printf("%d", arr[1]); return 0; }`,
      opts: [`10`, `20`, `30`, `Compiler error`],
      correct: 1,
    },
    // 6
    {
      q: `What does scanf() do in C?`,
      opts: [
        `Prints formatted output`,
        `Reads formatted input from stdin`,
        `Scans the file system`,
        `Allocates memory`,
      ],
      correct: 1,
    },
    // 7
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int add(int a, int b) { return a + b; }
int main() { printf("%d", add(3, 4)); return 0; }`,
      opts: [`3`, `4`, `7`, `Compiler error`],
      correct: 2,
    },
    // 8
    {
      q: `Which header file provides strlen()?`,
      opts: [`<stdio.h>`, `<stdlib.h>`, `<string.h>`, `<math.h>`],
      correct: 2,
    },
    // 9
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int x = 10; int *p = &x; printf("%d", *p); return 0; }`,
      opts: [`x`, `10`, `Address of x`, `Compiler error`],
      correct: 1,
    },
    // 10
    {
      q: `What is the size of an int on most 64-bit systems in C?`,
      opts: [`2 bytes`, `4 bytes`, `8 bytes`, `Depends on compiler`],
      correct: 1,
    },
    // 11
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int x = 5; x *= 3; printf("%d", x); return 0; }`,
      opts: [`8`, `15`, `3`, `Compiler error`],
      correct: 1,
    },
    // 12
    {
      q: `Which loop is guaranteed to execute its body at least once in C?`,
      opts: [`for`, `while`, `do-while`, `foreach`],
      correct: 2,
    },
    // 13
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int i = 0; while(i < 3) { printf("%d", i); i++; } return 0; }`,
      opts: [`123`, `012`, `0 1 2`, `Compiler error`],
      correct: 1,
    },
    // 14
    {
      q: `What does the break statement do inside a loop in C?`,
      opts: [
        `Restarts the loop`,
        `Skips the current iteration`,
        `Exits the loop immediately`,
        `Exits the program`,
      ],
      correct: 2,
    },
    // 15
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int a = 6, b = 4; printf("%d", a & b); return 0; }`,
      opts: [`2`, `4`, `6`, `10`],
      correct: 1,
    },
    // 16
    {
      q: `Which keyword is used to define a structure in C?`,
      opts: [`class`, `object`, `struct`, `record`],
      correct: 2,
    },
    // 17
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { char s[] = "Hello"; printf("%lu", sizeof(s)); return 0; }`,
      opts: [`5`, `6`, `Depends`, `Compiler error`],
      correct: 1,
    },
    // 18
    {
      q: `In C, what is the default return type of main()?`,
      opts: [`void`, `int`, `char`, `float`],
      correct: 1,
    },
    // 19
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() { int a = 2; printf("%d", a << 2); return 0; }`,
      opts: [`4`, `6`, `8`, `10`],
      correct: 2,
    },
    // 20
    {
      q: `What does the continue statement do inside a loop in C?`,
      opts: [
        `Exits the loop`,
        `Skips the rest of the current iteration and continues to the next`,
        `Restarts the program`,
        `Exits the function`,
      ],
      correct: 1,
    },
  ],

  hard: [
    // 1
    {
      q: `What does the static keyword do for a local variable in C?`,
      opts: [
        `Allocates it on the heap`,
        `Persists its value across function calls`,
        `Makes it visible globally`,
        `Prevents modification`,
      ],
      correct: 1,
    },
    // 2
    {
      q: `Which of the following leads to undefined behavior in C?`,
      opts: [
        `Dereferencing a NULL pointer`,
        `Using sizeof on an array`,
        `Declaring a struct`,
        `Including a header twice with include guards`,
      ],
      correct: 0,
    },
    // 3
    {
      q: `Which header provides memcpy() in C?`,
      opts: [`<stdio.h>`, `<string.h>`, `<stdlib.h>`, `<memory.h>`],
      correct: 1,
    },
    // 4
    {
      q: `Which of these is true about the expression i++ > ++i in C?`,
      opts: [
        `Well-defined; equals false`,
        `Undefined behaviour`,
        `Always evaluates to true`,
        `Compiler error`,
      ],
      correct: 1,
    },
    // 5
    {
      q: `What is the primary reason to declare a variable as volatile in C?`,
      opts: [
        `Prevent the compiler from optimizing away reads/writes it may see as redundant`,
        `Make the variable read-only`,
        `Force the variable into a register`,
        `Increase floating-point precision`,
      ],
      correct: 0,
    },
    // 6
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() {
  int a = 5;
  int *p = &a;
  *p = 10;
  printf("%d", a);
  return 0;
}`,
      opts: [`5`, `10`, `Address of a`, `Compiler error`],
      correct: 1,
    },
    // 7
    {
      q: `Which of the following correctly uses a function pointer in C?`,
      opts: [
        `int *fp(int);`,
        `int (*fp)(int);`,
        `*(int fp)(int);`,
        `int fp*(int);`,
      ],
      correct: 1,
    },
    // 8
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() {
  int arr[] = {1, 2, 3, 4, 5};
  int *p = arr;
  printf("%d", *(p + 2));
  return 0;
}`,
      opts: [`1`, `2`, `3`, `Compiler error`],
      correct: 2,
    },
    // 9
    {
      q: `What is a dangling pointer in C?`,
      opts: [
        `A pointer that has never been initialised`,
        `A pointer that points to freed or out-of-scope memory`,
        `A pointer to a global variable`,
        `A pointer used in a struct`,
      ],
      correct: 1,
    },
    // 10
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
void f(int *x) { *x = 20; }
int main() {
  int a = 10;
  f(&a);
  printf("%d", a);
  return 0;
}`,
      opts: [`10`, `20`, `0`, `Compiler error`],
      correct: 1,
    },
    // 11
    {
      q: `Which storage class makes a variable visible only within its translation unit in C?`,
      opts: [`extern`, `register`, `static`, `auto`],
      correct: 2,
    },
    // 12
    {
      q: `What is the result of sizeof(char) in C, guaranteed by the standard?`,
      opts: [`0`, `1`, `2`, `Platform-dependent`],
      correct: 1,
    },
    // 13
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() {
  int x = 0;
  x = x++ + ++x;
  printf("%d", x);
  return 0;
}`,
      opts: [`1`, `2`, `Undefined behaviour`, `Compiler error`],
      correct: 2,
    },
    // 14
    {
      q: `Which function releases dynamically allocated memory in C?`,
      opts: [`delete()`, `release()`, `free()`, `dealloc()`],
      correct: 2,
    },
    // 15
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() {
  int a = 5, b = 10;
  int *p = &a, *q = &b;
  *p = *q;
  printf("%d %d", a, b);
  return 0;
}`,
      opts: [`5 10`, `10 10`, `5 5`, `Compiler error`],
      correct: 1,
    },
    // 16
    {
      q: `What does the restrict keyword tell the compiler in C99?`,
      opts: [
        `The pointer is read-only`,
        `No other pointer will alias the same memory during its scope`,
        `The variable is thread-local`,
        `The memory is allocated on the stack`,
      ],
      correct: 1,
    },
    // 17
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
static int count = 0;
void inc() { count++; }
int main() { inc(); inc(); inc(); printf("%d", count); return 0; }`,
      opts: [`0`, `1`, `3`, `Compiler error`],
      correct: 2,
    },
    // 18
    {
      q: `Which of the following correctly defines a typedef for a struct Point in C?`,
      opts: [
        `typedef struct { int x, y; } Point;`,
        `define struct Point { int x, y; };`,
        `struct typedef Point { int x, y; };`,
        `type Point = struct { int x, y; };`,
      ],
      correct: 0,
    },
    // 19
    {
      q: `What is the output of the following C code?`,
      code: `#include <stdio.h>
int main() {
  char *s = "Hello";
  printf("%c", *(s + 1));
  return 0;
}`,
      opts: [`H`, `e`, `l`, `Compiler error`],
      correct: 1,
    },
    // 20
    {
      q: `What is the difference between calloc() and malloc() in C?`,
      opts: [
        `calloc allocates on stack; malloc on heap`,
        `calloc initialises allocated memory to zero; malloc does not`,
        `malloc takes two arguments; calloc takes one`,
        `There is no difference`,
      ],
      correct: 1,
    },
  ],
};

// C_POOL keys already match TECH_LEVEL_ORDER ("easy", "medium", "hard"),
// so map them directly — each tier keeps its 20 questions intact.
const C_QUESTIONS = Object.values(C_POOL).flatMap((arr) => arr);
const C_LEVEL_POOL = {
  easy: C_POOL.easy,
  medium: C_POOL.medium,
  hard: C_POOL.hard,
};

// ─── Assignment question bank: C Programming (50 Qs) ─────────────────────────
// Source: SECTION 1 – SECTION 5 handout. Covers Syntax & Basics, Control Flow,
// Pointers & Memory, Arrays & Strings, Structures / Preprocessors & Advanced.
const ASSIGNMENT_C_QUESTIONS = [
  // ── SECTION 1: SYNTAX & BASICS (Q1–Q10) ─────────────────────────────────
  {
    q: "Which keyword is used to prevent any changes to a variable in C?",
    opts: ["volatile", "immutable", "const", "stable"],
    correct: 2,
  },
  {
    q: "What is the default return type of a function in C if it is not explicitly specified?",
    opts: ["void", "int", "char", "float"],
    correct: 1,
  },
  {
    q: "Which of the following is NOT a valid variable name in C?",
    opts: ["_int", "variable_1", "2variable", "main"],
    correct: 2,
  },
  {
    q: "What is the size of a char data type in standard C?",
    opts: ["1 byte", "2 bytes", "4 bytes", "Depends on the compiler"],
    correct: 0,
  },
  {
    q: "Which operator is used to find the remainder of a division between two integers?",
    opts: ["/", "%", "\\", "rem"],
    correct: 1,
  },
  {
    q: "Which escape sequence represents a horizontal tab in C?",
    opts: ["\\b", "\\n", "\\t", "\\r"],
    correct: 2,
  },
  {
    q: "What is the correct way to write a single-line comment in C99 and later?",
    opts: ["# Comment", "// Comment", "/* Comment", "-- Comment"],
    correct: 1,
  },
  {
    q: "Which format specifier is used to print a double precision floating-point number using printf?",
    opts: ["%f", "%lf", "%d", "%ld"],
    correct: 0,
  },
  {
    q: "What is the result of the expression 5 / 2 in C?",
    opts: ["2.5", "2", "3", "0"],
    correct: 1,
  },
  {
    q: "Which standard library header file contains the printf() and scanf() functions?",
    opts: ["stdlib.h", "conio.h", "string.h", "stdio.h"],
    correct: 3,
  },
  // ── SECTION 2: CONTROL FLOW & LOOPS (Q11–Q20) ───────────────────────────
  {
    q: "Which block always executes at least once in a C program?",
    opts: ["while loop", "for loop", "do-while loop", "if-else statement"],
    correct: 2,
  },
  {
    q: "What happens when a break statement is encountered inside a nested loop?",
    opts: [
      "It terminates the innermost loop.",
      "It terminates all loops.",
      "It skips the current iteration.",
      "It exits the entire program.",
    ],
    correct: 0,
  },
  {
    q: "What is the purpose of the continue statement in a loop?",
    opts: [
      "To terminate the loop entirely.",
      "To skip the rest of the current iteration and move to the next.",
      "To restart the loop from the beginning variable values.",
      "To pause execution until user input.",
    ],
    correct: 1,
  },
  {
    q: "What are the only values evaluated as 'false' in a C conditional statement?",
    opts: ["Only -1", "Only 0", "Any negative number", "NULL and 0"],
    correct: 1,
  },
  {
    q: "Can a switch statement evaluate a floating-point expression?",
    opts: [
      "Yes, always",
      "Yes, if cast to a double",
      "No, only integer and character types are allowed",
      "Only in C11 standard",
    ],
    correct: 2,
  },
  {
    q: "What is the equivalent ternary operator structure for an if-else statement?",
    opts: [
      "condition ? expression1 : expression2",
      "condition : expression1 ? expression2",
      "condition ? expression1 ? expression2",
      "expression1 ? condition : expression2",
    ],
    correct: 0,
  },
  {
    q: "What type of loop is for(;;)?",
    opts: [
      "Syntax error loop",
      "Finite loop",
      "Infinite loop",
      "Undefined loop",
    ],
    correct: 2,
  },
  {
    q: "What is the order of evaluation in a for loop header?",
    opts: [
      "Condition, Initialization, Increment",
      "Initialization, Condition, Increment",
      "Increment, Condition, Initialization",
      "Initialization, Increment, Condition",
    ],
    correct: 1,
  },
  {
    q: "Which header file is required to use the exit() function to stop program execution?",
    opts: ["stdio.h", "stdlib.h", "process.h", "assert.h"],
    correct: 1,
  },
  {
    q: 'What is a "dangling else" problem in C?',
    opts: [
      "An else statement without an executable body",
      "An else statement that cannot be matched to any if statement",
      "Nested if-else statements without braces causing ambiguity about which if matches the else",
      "Memory leaks caused by conditions",
    ],
    correct: 2,
  },
  // ── SECTION 3: POINTERS & MEMORY MANAGEMENT (Q21–Q30) ───────────────────
  {
    q: "What does a pointer variable store?",
    opts: [
      "The value of another variable",
      "The memory address of another variable",
      "The name of another variable",
      "A character string",
    ],
    correct: 1,
  },
  {
    q: "Which operator is used to obtain the address of a variable?",
    opts: ["*", "->", "&", "."],
    correct: 2,
  },
  {
    q: "What is a NULL pointer?",
    opts: [
      "A pointer pointing to the address 0xFFFF",
      "A pointer that has not been initialized",
      "A pointer pointing to a reliable empty string",
      "A pointer that points to nothing or address 0",
    ],
    correct: 3,
  },
  {
    q: "Which function is used to allocate memory dynamically but initializes it all to zero?",
    opts: ["malloc()", "calloc()", "realloc()", "free()"],
    correct: 1,
  },
  {
    q: "What happens if you try to free dynamically allocated memory twice?",
    opts: [
      "Nothing, it is safe.",
      "The memory is reallocated.",
      "It causes undefined behavior or a runtime crash.",
      "The pointer is automatically set to NULL.",
    ],
    correct: 2,
  },
  {
    q: 'What is a "void pointer" in C?',
    opts: [
      "A pointer that cannot point to anything",
      "A pointer that points to a function returning void",
      "A generic pointer that can point to any data type",
      "A pointer that points to deleted memory",
    ],
    correct: 2,
  },
  {
    q: "If ptr is a pointer to an integer, what does ptr++ do?",
    opts: [
      "Increments the integer value stored at that address by 1",
      "Moves the pointer to the next memory address based on the size of an int",
      "Adds 1 byte to the pointer address",
      "Multiplies the address by 2",
    ],
    correct: 1,
  },
  {
    q: "Which function changes the size of previously allocated memory without losing current data?",
    opts: ["malloc()", "calloc()", "realloc()", "alloc()"],
    correct: 2,
  },
  {
    q: 'What is a "dangling pointer"?',
    opts: [
      "A pointer pointing to an unallocated memory slot in a loop",
      "A pointer pointing to a memory location that has already been deallocated",
      "A pointer that holds an integer value instead of an address",
      "A pointer initialized to NULL",
    ],
    correct: 1,
  },
  {
    q: "What operator is used to access structure members using a structure pointer?",
    opts: [".", "*", "&", "->"],
    correct: 3,
  },
  // ── SECTION 4: ARRAYS & STRINGS (Q31–Q40) ───────────────────────────────
  {
    q: "What is the index of the first element in a C array?",
    opts: ["-1", "0", "1", "Depends on the declaration"],
    correct: 1,
  },
  {
    q: "What character automatically marks the end of a string array in C?",
    opts: ["\\n", "\\0", "\\t", "EOF"],
    correct: 1,
  },
  {
    q: "What happens if you access an array index out of its declared bounds?",
    opts: [
      "Compiler error",
      "Runtime exception thrown by OS",
      "Undefined behavior (could read garbage value or crash)",
      "The array automatically resizes",
    ],
    correct: 2,
  },
  {
    q: "Which library function is used to compare two strings alphabetically?",
    opts: ["strcpy()", "strcmp()", "strlen()", "strcat()"],
    correct: 1,
  },
  {
    q: "How is a multi-dimensional array stored sequentially in memory in C?",
    opts: [
      "Column-major order",
      "Row-major order",
      "Zig-zag order",
      "It is non-contiguous",
    ],
    correct: 1,
  },
  {
    q: "What does strlen() count when calculating string length?",
    opts: [
      "All characters including the null terminator",
      "All characters excluding the null terminator",
      "Only alphabetic characters",
      "Total bytes allocated for the array",
    ],
    correct: 1,
  },
  {
    q: "If you declare int arr[5] = {1, 2};, what are the values of the remaining elements?",
    opts: ["Garbage values", "1, 2 repeated", "0, 0, 0", "Compiler error"],
    correct: 2,
  },
  {
    q: "Which function safely limits the number of characters read into a string from standard input?",
    opts: ["gets()", "fgets()", 'scanf("%s")', "getstr()"],
    correct: 1,
  },
  {
    q: "What does the array name itself represent when passed to a function?",
    opts: [
      "The value of the first element",
      "The size of the array",
      "A pointer to the first element of the array",
      "The total number of elements",
    ],
    correct: 2,
  },
  {
    q: "Which function copies one string to another destination memory block?",
    opts: ["strdup()", "strcat()", "strcpy()", "strchr()"],
    correct: 2,
  },
  // ── SECTION 5: STRUCTURES, PREPROCESSORS & ADVANCED (Q41–Q50) ───────────
  {
    q: "What is the main structural difference between a struct and a union in C?",
    opts: [
      "Structs can hold functions, unions cannot.",
      "Struct elements share the same memory space; union elements have separate spaces.",
      "Union elements share the same memory space; struct elements have separate spaces.",
      "Unions cannot hold pointer types.",
    ],
    correct: 2,
  },
  {
    q: "Which preprocessor directive is used to define constants or macros?",
    opts: ["#include", "#define", "#ifdef", "#pragma"],
    correct: 1,
  },
  {
    q: "What does the #include <filename> syntax indicate to the preprocessor?",
    opts: [
      "Search for the file only in the current working directory.",
      "Search for the file in the standard system include directories.",
      "Download the file from an external repository.",
      "Compile the file separately.",
    ],
    correct: 1,
  },
  {
    q: "Which storage class specifies that a variable should be kept in a CPU register if possible?",
    opts: ["auto", "static", "register", "extern"],
    correct: 2,
  },
  {
    q: "What is the lifetime of a local static variable inside a function?",
    opts: [
      "Destroys when the function exits",
      "Destroys when the current loop block exits",
      "Persists throughout the entire lifetime of the program",
      "Persists only until the next function call",
    ],
    correct: 2,
  },
  {
    q: "What does the extern keyword signify?",
    opts: [
      "The variable is visible only within the local function.",
      "The variable is declared in another file or outside the current scope.",
      "The variable is stored in external flash memory.",
      "The variable cannot be altered during execution.",
    ],
    correct: 1,
  },
  {
    q: "Which file opening mode in fopen() is used to open a file for writing at the end of it?",
    opts: ['"r"', '"w"', '"a"', '"r+"'],
    correct: 2,
  },
  {
    q: "What is the purpose of a typedef in C?",
    opts: [
      "To declare a new data type entirely",
      "To create an alias or shortcut name for an existing data type",
      "To convert variables from one type to another dynamically",
      "To encrypt variable names",
    ],
    correct: 1,
  },
  {
    q: "What is a macro expansion?",
    opts: [
      "A runtime optimization tool",
      "A compile-time textual substitution text process",
      "A dynamic memory expansion algorithm",
      "A mechanism to create deep structures",
    ],
    correct: 1,
  },
  {
    q: "Which operator is used to perform bitwise AND operations in C?",
    opts: ["&&", "&", "|", "^"],
    correct: 1,
  },
];

const C_POOL_FLAT = TECH_LEVEL_ORDER.flatMap((level) => C_LEVEL_POOL[level]);

const JAVA_QUESTIONS = [
    { q: `Which statement about HashMap is correct?`, code: ``, opts: [`It is synchronized by default`,`It permits one null key and multiple null values`,`It sorts keys automatically`,`It rejects duplicate values`], correct: 1 },
    { q: `What is the result of Integer a=128; Integer b=128; a==b?`, code: `Integer a = 128;\nInteger b = 128;\nSystem.out.println(a == b);`, opts: [`Always true`,`Always false`,`Usually false because of Integer caching range`,`Compilation error`], correct: 2 },
    { q: `Which keyword guarantees visibility but not atomicity?`, code: ``, opts: [`transient`,`volatile`,`native`,`strictfp`], correct: 1 },
    { q: `Which collection provides O(log n) search based on sorted keys?`, code: ``, opts: [`HashMap`,`TreeMap`,`IdentityHashMap`,`WeakHashMap`], correct: 1 },
    { q: `What does try-with-resources require?`, code: `try (java.io.BufferedReader r = new java.io.BufferedReader(new java.io.FileReader("file.txt"))) { String s = r.readLine(); }`, opts: [`Cloneable`,`Serializable`,`AutoCloseable`,`Runnable`], correct: 2 },
    { q: `Which is true about final fields?`, code: ``, opts: [`They can never be initialized`,`They receive special safe-publication guarantees after construction`,`They are always immutable`,`They are always static`], correct: 1 },
    { q: `Which method is used to wait for a thread to finish?`, code: `Thread t = new Thread(() -> {});\nt.start();\nt.join();`, opts: [`sleep()`,`wait()`,`join()`,`yield()`], correct: 2 },
    { q: `What does ConcurrentHashMap disallow?`, code: ``, opts: [`Concurrent reads`,`Null keys and null values`,`Atomic updates`,`Weak consistency`], correct: 1 },
    { q: `Which operation is terminal in a Stream?`, code: `java.util.List<Integer> l = java.util.stream.Stream.of(1,2,3).map(i -> i*2).collect(java.util.stream.Collectors.toList());`, opts: [`map`,`filter`,`sorted`,`collect`], correct: 3 },
    { q: `What does PECS mean?`, code: ``, opts: [`Public Encapsulation Class System`,`Producer Extends, Consumer Super`,`Private Exception Control Syntax`,`Parallel Execution Collection Strategy`], correct: 1 },
    { q: `Which class is immutable?`, code: ``, opts: [`StringBuilder`,`ArrayList`,`String`,`Date`], correct: 2 },
    { q: `What is method overloading resolved primarily at?`, code: ``, opts: [`Runtime`,`Compile time`,`Garbage collection`,`Class unloading`], correct: 1 },
    { q: `Which exception is unchecked?`, code: ``, opts: [`IOException`,`SQLException`,`ClassNotFoundException`,`NullPointerException`], correct: 3 },
    { q: `What does a synchronized instance method lock?`, code: ``, opts: [`Class object`,`Current instance`,`Method object`,`Thread object`], correct: 1 },
    { q: `Which interface represents a task returning a result?`, code: ``, opts: [`Runnable`,`Callable`,`Consumer`,`Supplier only`], correct: 1 },
    { q: `What is the main purpose of equals() and hashCode() consistency?`, code: ``, opts: [`Faster compilation`,`Correct hashed-collection behavior`,`Prevent inheritance`,`Enable serialization`], correct: 1 },
    { q: `Which map maintains insertion order?`, code: ``, opts: [`HashMap`,`TreeMap`,`LinkedHashMap`,`Hashtable`], correct: 2 },
    { q: `What does Optional.orElseGet accept?`, code: ``, opts: [`A value`,`A Supplier`,`A Predicate`,`A Consumer`], correct: 1 },
    { q: `Which garbage collector targets very low pause times?`, code: ``, opts: [`Serial only`,`ZGC`,`Epsilon only`,`CMS exclusively`], correct: 1 },
    { q: `What is a deadlock?`, code: ``, opts: [`A fast computation`,`Threads permanently waiting for each other`,`A completed future`,`A compiler warning`], correct: 1 },
    { q: `Which is true about interfaces?`, code: ``, opts: [`They cannot contain default methods`,`They can have abstract and default methods`,`They must have constructors`,`They cannot be generic`], correct: 1 },
    { q: `What does flatMap primarily do?`, code: ``, opts: [`Sort values`,`Flatten nested streams`,`Cache results`,`Create threads`], correct: 1 },
    { q: `Which reference type does not prevent garbage collection and is commonly used for caches?`, code: ``, opts: [`Strong`,`Weak`,`Final`,`Constant`], correct: 1 },
    { q: `What is type erasure?`, code: ``, opts: [`Removing comments`,`Removing generic type information largely at runtime`,`Deleting classes`,`Compressing bytecode`], correct: 1 },
    { q: `Which lock supports tryLock()?`, code: `java.util.concurrent.locks.ReentrantLock lock = new java.util.concurrent.locks.ReentrantLock();\nif (lock.tryLock()) { try { /*...*/ } finally { lock.unlock(); } }`, opts: [`ReentrantLock`,`Object only`,`String`,`ThreadLocal`], correct: 0 },
    { q: `What does volatile not guarantee?`, code: ``, opts: [`Visibility`,`Ordering constraints`,`Atomic compound operations`,`Reads seeing recent writes`], correct: 2 },
    { q: `Which stream operation may be stateful and expensive?`, code: ``, opts: [`filter`,`sorted`,`map`,`peek`], correct: 1 },
    { q: `What is a record primarily designed for?`, code: ``, opts: [`Mutable GUI state`,`Concise data carriers`,`Thread scheduling`,`Bytecode generation`], correct: 1 },
    { q: `Which class loader normally delegates to its parent first?`, code: ``, opts: [`Bootstrap/class-loading delegation model`,`Every custom loader never delegates`,`Only URLClassLoader after Java 1`,`None`], correct: 0 },
    { q: `What happens when a checked exception is not caught or declared?`, code: ``, opts: [`Runtime ignores it`,`Compilation fails`,`JVM restarts`,`It becomes unchecked`], correct: 1 },
    { q: `Which collector characteristic helps a parallel collector combine results?`, code: ``, opts: [`combiner`,`destructor`,`finalizer`,`serializer`], correct: 0 },
    { q: `What does String.intern() relate to?`, code: `String s = new String("x"); String t = s.intern();`, opts: [`Thread pool`,`String pool`,`Heap compaction only`,`File descriptors`], correct: 1 },
    { q: `Which is safest for lazy singleton initialization?`, code: ``, opts: [`Unsynchronized static field`,`Initialization-on-demand holder idiom`,`Public mutable instance`,`Repeated new calls`], correct: 1 },
    { q: `What does a Spliterator support?`, code: ``, opts: [`Controlled traversal and splitting`,`Serialization only`,`Database locking`,`Exception suppression`], correct: 0 },
    { q: `Which operation short-circuits?`, code: ``, opts: [`map`,`anyMatch`,`sorted`,`collect`], correct: 1 },
    { q: `What is a bridge method?`, code: ``, opts: [`Network method`,`Compiler-generated method preserving polymorphism with generics`,`Native method`,`Deprecated method`], correct: 1 },
    { q: `Which executor is suited to scheduled tasks?`, code: `java.util.concurrent.ScheduledExecutorService s = java.util.concurrent.Executors.newScheduledThreadPool(1);`, opts: [`ScheduledExecutorService`,`ForkJoinTask only`,`ErrorHandler`,`ClassLoader`], correct: 0 },
    { q: `What does synchronized provide?`, code: ``, opts: [`Mutual exclusion and visibility guarantees`,`Only faster execution`,`Only inheritance`,`Automatic immutability`], correct: 0 },
    { q: `Which statement about records is correct?`, code: ``, opts: [`Records can extend arbitrary classes`,`Records implicitly extend Record`,`Records must be mutable`,`Records cannot implement interfaces`], correct: 1 },
    { q: `Which operation can throw ConcurrentModificationException with ArrayList?`, code: `List<Integer> a = new ArrayList<>(); for (Integer x : a) { a.add(5); }`, opts: [`Structural modification during fail-fast iteration`,`Reading size()` ,`Calling get()`,`Creating an iterator`], correct: 0 },
    { q: `What is the purpose of a defensive copy?`, code: ``, opts: [`Protect internal mutable state`,`Increase inheritance`,`Disable GC`,`Avoid compilation`], correct: 0 },
    { q: `Which annotation indicates an intended functional interface?`, code: `@FunctionalInterface\npublic interface Fn { void apply(); }`, opts: [`@Override`,`@FunctionalInterface`,`@SafeVarargs only`,`@Immutable`], correct: 1 },
    { q: `What does CompletableFuture.exceptionally() do?`, code: `CompletableFuture.supplyAsync(() -> 1).exceptionally(ex -> 0);`, opts: [`Handles an exceptional completion`,`Starts a thread necessarily`,`Closes resources`,`Serializes a result`], correct: 0 },
    { q: `Which data structure is lock-free for many atomic numeric updates?`, code: `java.util.concurrent.atomic.LongAdder a = new java.util.concurrent.atomic.LongAdder(); a.increment();`, opts: [`LongAdder`,`String`,`TreeSet`,`Formatter`], correct: 0 },
    { q: `What is class identity determined by?`, code: ``, opts: [`Class name only`,`Class name plus defining class loader`,`Package only`,`Source file name only`], correct: 1 },
    { q: `Which keyword prevents a method from being overridden?`, code: ``, opts: [`transient`,`final`,`volatile`,`synchronized`], correct: 1 },
    { q: `What does map.computeIfAbsent do?`, code: `map.computeIfAbsent(key, k -> new ArrayList<>());`, opts: [`Computes and associates a value when key is absent`,`Deletes all keys`,`Sorts a map`,`Clones a map`], correct: 0 },
    { q: `Which is true of parallel streams?`, code: ``, opts: [`They always improve performance`,`They may hurt performance for small or blocking workloads`,`They are always deterministic in ordering`,`They require synchronized collections`], correct: 1 },
    { q: `What is a sealed class used for?`, code: ``, opts: [`Restricting permitted subclasses`,`Preventing object creation always`,`Enabling native code`,`Disabling pattern matching`], correct: 0 },
    { q: `Which I/O API is best suited to asynchronous file operations?`, code: ``, opts: [`AsynchronousFileChannel`,`File only`,`PrintStream only`,`Scanner only`], correct: 0 },
    { q: `What does readObject() customize?`, code: `private void readObject(java.io.ObjectInputStream in) throws Exception { in.defaultReadObject(); }`, opts: [`Java serialization deserialization`,`Stream filtering`,`Thread creation`,`Hashing only`], correct: 0 },
    { q: `Which statement about virtual threads is generally correct?`, code: ``, opts: [`They are intended for many blocking tasks`,`They require one OS thread each`,`They replace all CPU parallelism`,`They cannot use executors`], correct: 0 },
    { q: `What does escape analysis potentially enable?`, code: ``, opts: [`Lock elimination and scalar replacement`,`More checked exceptions`,`Runtime inheritance`,`Source formatting`], correct: 0 },
    { q: `Which map compares keys using == rather than equals()?`, code: `Map<Object,String> m = new IdentityHashMap<>();`, opts: [`TreeMap`,`IdentityHashMap`,`LinkedHashMap`,`EnumMap`], correct: 1 },
    { q: `What is the identity element in reduce required to be?`, code: ``, opts: [`Neutral for the accumulator`,`Always null`,`A thread`,`A checked exception`], correct: 0 },
    { q: `Which feature enables exhaustive handling of restricted hierarchies?`, code: ``, opts: [`Sealed types with pattern matching`,`transient fields`,`native methods`,`assertions only`], correct: 0 },
    { q: `What is a phantom reference mainly used with?`, code: ``, opts: [`Post-mortem cleanup tracking`,`Direct object access`,`String concatenation`,`Method overloading`], correct: 0 },
    { q: `Which operation on a parallel stream must have an associative accumulator for correctness?`, code: ``, opts: [`reduce`,`peek`,`limit only`,`iterator creation`], correct: 0 },
    { q: `What does ReentrantReadWriteLock provide?`, code: `java.util.concurrent.locks.ReentrantReadWriteLock rw = new java.util.concurrent.locks.ReentrantReadWriteLock();`, opts: [`Separate read and write locks`,`Only a semaphore`,`Immutable objects`,`Automatic transactions`], correct: 0 },
    { q: `What does Files.lines() return?`, code: `try (java.util.stream.Stream<String> s = java.nio.file.Files.lines(java.nio.file.Path.of("file.txt"))) { s.forEach(System.out::println); }`, opts: [`Lazy Stream of lines`,`Eager List always`,`ByteBuffer only`,`Thread pool`], correct: 0 },
];

const QUESTIONS = {
  aptitude: {
    label: "Aptitude",
    code: "APT-01",
    items: [
      {
        q: "A train travels 60 km in 45 minutes. What is its speed in km/h?",
        opts: ["60", "75", "80", "90"],
        correct: 2,
      },
      {
        q: "A sum of money doubles itself in 8 years at simple interest. What is the rate of interest per annum?",
        opts: ["10%", "12.5%", "15%", "8%"],
        correct: 1,
      },
      {
        q: "Find the next number: 2, 6, 12, 20, 30, ?",
        opts: ["40", "42", "44", "36"],
        correct: 1,
      },
      {
        q: "The average of 5 numbers is 20. Excluding one number, the average becomes 15. Find the excluded number.",
        opts: ["30", "35", "40", "45"],
        correct: 2,
      },
      {
        q: "A can finish a job in 10 days, B in 15 days. Working together, how many days will it take?",
        opts: ["5", "6", "7", "8"],
        correct: 1,
      },
      { q: "What is 15% of 240?", opts: ["30", "36", "32", "40"], correct: 1 },
      {
        q: "Two numbers are in ratio 3:4 and sum to 63. Find the larger number.",
        opts: ["27", "36", "32", "28"],
        correct: 1,
      },
      {
        q: "A shopkeeper marks an item 25% above cost and gives a 10% discount. Find his profit %.",
        opts: ["12.5%", "15%", "10%", "17.5%"],
        correct: 0,
      },
      {
        q: "20 men can build a wall in 15 days. How many days will 30 men take?",
        opts: ["8", "10", "12", "15"],
        correct: 1,
      },
      {
        q: "Find the compound interest on Rs. 5000 at 10% p.a. for 2 years.",
        opts: ["1000", "1050", "1100", "1150"],
        correct: 1,
      },
      {
        q: "Find the HCF of 18 and 24.",
        opts: ["4", "6", "8", "12"],
        correct: 1,
      },
      {
        q: "Find the LCM of 4 and 6.",
        opts: ["8", "10", "12", "24"],
        correct: 2,
      },
      {
        q: "A man's age is 3 times his son's age. In 15 years he will be twice his son's age. Find the son's current age.",
        opts: ["10", "15", "20", "12"],
        correct: 1,
      },
      {
        q: "What is the probability of getting a head when a fair coin is tossed once?",
        opts: ["1/4", "1/3", "1/2", "1"],
        correct: 2,
      },
      {
        q: "What is the probability of drawing a king from a standard deck of 52 cards?",
        opts: ["1/13", "1/26", "1/4", "1/52"],
        correct: 0,
      },
      {
        q: "In how many ways can the letters of the word 'CAT' be arranged?",
        opts: ["3", "6", "9", "12"],
        correct: 1,
      },
      {
        q: "If x:y = 2:3 and y:z = 4:5, find x:z.",
        opts: ["8:15", "2:5", "6:15", "4:9"],
        correct: 0,
      },
      {
        q: "A boat covers 10 km upstream in 2 hours and the same distance downstream in 1 hour. Find the boat's speed in still water.",
        opts: ["6.5 km/h", "7 km/h", "7.5 km/h", "8 km/h"],
        correct: 2,
      },
      {
        q: "The cost price of 20 articles equals the selling price of 16 articles. Find the profit percentage.",
        opts: ["16%", "20%", "25%", "30%"],
        correct: 2,
      },
      {
        q: "A alone can do a job in 12 days, B alone in 18 days. In how many days will they finish it together?",
        opts: ["6", "6.5", "7.2", "8"],
        correct: 2,
      },
      {
        q: "Simplify: 3/4 + 1/6",
        opts: ["5/6", "7/12", "11/12", "4/10"],
        correct: 2,
      },
      {
        q: "What is 25% of 25% of 800?",
        opts: ["40", "50", "60", "80"],
        correct: 1,
      },
      {
        q: "A sum of Rs. 2000 amounts to Rs. 2420 in 2 years at compound interest. Find the rate.",
        opts: ["8%", "10%", "12%", "15%"],
        correct: 1,
      },
      {
        q: "Find the odd one out: 3, 5, 7, 9, 11",
        opts: ["3", "9", "7", "11"],
        correct: 1,
      },
      {
        q: "If a:b = 3:4 and b:c = 6:7, find a:b:c.",
        opts: ["9:12:14", "3:4:7", "6:8:14", "9:12:16"],
        correct: 0,
      },
      {
        q: "Pipe A fills a tank in 20 minutes, pipe B in 30 minutes. How long will both take together?",
        opts: ["10 min", "12 min", "15 min", "18 min"],
        correct: 1,
      },
      {
        q: "Find the next number in the series: 1, 1, 2, 3, 5, 8, ?",
        opts: ["11", "12", "13", "14"],
        correct: 2,
      },
      {
        q: "Find the perimeter of a square with side 12 cm.",
        opts: ["36 cm", "42 cm", "48 cm", "54 cm"],
        correct: 2,
      },
      {
        q: "Find the simple interest on Rs. 1500 at 6% per annum for 4 years.",
        opts: ["300", "330", "360", "400"],
        correct: 2,
      },
      {
        q: "A car covers 240 km in 4 hours. What is its average speed?",
        opts: ["50 km/h", "55 km/h", "60 km/h", "65 km/h"],
        correct: 2,
      },
    ],
  },
  verbal: {
    label: "Verbal",
    code: "VRB-01",
    items: [
      {
        q: "Choose the synonym of 'Benevolent'.",
        opts: ["Kind", "Cruel", "Timid", "Greedy"],
        correct: 0,
      },
      {
        q: "Choose the antonym of 'Meticulous'.",
        opts: ["Careless", "Precise", "Thorough", "Careful"],
        correct: 0,
      },
      {
        q: "She has a ____ for painting since childhood.",
        opts: ["passion", "passed", "passing", "passable"],
        correct: 0,
      },
      {
        q: "Choose the correctly spelled word.",
        opts: ["Accomodate", "Acommodate", "Accommodate", "Accomadate"],
        correct: 2,
      },
      {
        q: "Identify the correct sentence.",
        opts: [
          "Neither of the boys were present",
          "Neither of the boys was present",
          "Neither of the boy were present",
          "Neither of boys was present",
        ],
        correct: 1,
      },
      {
        q: "One-word substitution for 'a person who can speak many languages'.",
        opts: ["Linguist", "Polyglot", "Interpreter", "Bilingual"],
        correct: 1,
      },
      {
        q: "Which idiom means 'to reveal a secret'?",
        opts: [
          "Let the cat out of the bag",
          "Beat around the bush",
          "Hit the nail",
          "Cut corners",
        ],
        correct: 0,
      },
      {
        q: "Complete the analogy: Doctor : Hospital :: Teacher : ?",
        opts: ["Book", "School", "Student", "Chalk"],
        correct: 1,
      },
      {
        q: "Choose the synonym of 'Candid'.",
        opts: ["Frank", "Deceptive", "Rude", "Hesitant"],
        correct: 0,
      },
      {
        q: "Choose the antonym of 'Abundant'.",
        opts: ["Scarce", "Plentiful", "Ample", "Rich"],
        correct: 0,
      },
      {
        q: "Choose the correctly spelled word.",
        opts: ["Occassion", "Ocassion", "Occasion", "Occasoin"],
        correct: 2,
      },
      {
        q: "He is allergic ____ peanuts.",
        opts: ["to", "from", "of", "with"],
        correct: 0,
      },
      {
        q: "One-word substitution for 'one who loves books'.",
        opts: ["Bibliophile", "Bibliomaniac", "Librarian", "Novelist"],
        correct: 0,
      },
      {
        q: "Which idiom means 'to postpone something'?",
        opts: ["Put off", "Put up", "Put on", "Put down"],
        correct: 0,
      },
      {
        q: "Complete the analogy: Pen : Write :: Knife : ?",
        opts: ["Cut", "Sharp", "Kitchen", "Blade"],
        correct: 0,
      },
      {
        q: "Identify the correct sentence.",
        opts: [
          "He don't like tea",
          "He doesn't like tea",
          "He not like tea",
          "He not likes tea",
        ],
        correct: 1,
      },
      {
        q: "Choose the synonym of 'Ambiguous'.",
        opts: ["Unclear", "Obvious", "Certain", "Direct"],
        correct: 0,
      },
      {
        q: "Choose the antonym of 'Genuine'.",
        opts: ["Fake", "Real", "Authentic", "True"],
        correct: 0,
      },
      {
        q: "She is good ____ mathematics.",
        opts: ["at", "in", "on", "with"],
        correct: 0,
      },
      {
        q: "Choose the correctly spelled word.",
        opts: ["Definately", "Definitely", "Definitly", "Definetely"],
        correct: 1,
      },
      {
        q: "One-word substitution for 'a place where books are kept for reading'.",
        opts: ["Library", "Bookstore", "Archive", "Museum"],
        correct: 0,
      },
      {
        q: "Which idiom means 'to reveal a secret unintentionally'?",
        opts: [
          "Spill the beans",
          "Cook the books",
          "Miss the boat",
          "Kill time",
        ],
        correct: 0,
      },
      {
        q: "Complete the analogy: Bird : Nest :: Bee : ?",
        opts: ["Hive", "Web", "Nest", "Cave"],
        correct: 0,
      },
      {
        q: "Identify the correct sentence.",
        opts: [
          "Each of the students have submitted their assignment",
          "Each of the students has submitted their assignment",
          "Each of the student have submitted assignment",
          "Each of student has submit assignment",
        ],
        correct: 1,
      },
      {
        q: "Choose the synonym of 'Diligent'.",
        opts: ["Hardworking", "Lazy", "Careless", "Slow"],
        correct: 0,
      },
      {
        q: "Choose the antonym of 'Optimistic'.",
        opts: ["Pessimistic", "Hopeful", "Positive", "Cheerful"],
        correct: 0,
      },
      {
        q: "The committee ____ decided to postpone the meeting.",
        opts: ["has", "have", "having", "had"],
        correct: 0,
      },
      {
        q: "One-word substitution for 'a person who does not believe in the existence of God'.",
        opts: ["Atheist", "Agnostic", "Theist", "Pagan"],
        correct: 0,
      },
      {
        q: "Which idiom means 'to work late into the night'?",
        opts: [
          "Burn the midnight oil",
          "Waste time",
          "Cook dinner",
          "Take a break",
        ],
        correct: 0,
      },
      {
        q: "Complete the analogy: Author : Book :: Sculptor : ?",
        opts: ["Statue", "Chisel", "Museum", "Gallery"],
        correct: 0,
      },
    ],
  },
  technical: {
    label: "Technical · Python",
    code: "TCH-PY-01",
    tiers: "Easy → Medium → Hard",
    quizCount: 60,
    poolSize: 250,
    get items() {
      return TECH_POOL_FLAT;
    },
  },
  technical_c: {
    label: "Technical · C",
    code: "TCH-C-01",
    tiers: "Easy → Medium → Hard",
    quizCount: 60,
    poolSize: C_QUESTIONS.length,
    get items() {
      return C_POOL_FLAT;
    },
  },
  technical_java: {
    label: "Technical · Java",
    code: "TCH-JAVA-01",
    tiers: "Single → Exam",
    quizCount: 60,
    poolSize: JAVA_QUESTIONS.length,
    items: JAVA_QUESTIONS,
  },
  assignment_c: {
    label: "C Programming — Assignment",
    code: "ASGN-C-01",
    description: "Exam starts at 5pm",
    items: ASSIGNMENT_C_QUESTIONS,
    isAssignment: true,
  },
};

const REVEAL_MS = 0;

function useStorage() {
  const [ready, setReady] = useState(false);
  useEffect(() => setReady(true), []);
  const get = async (key, shared = false) => {
    try {
      const r = await window.storage.get(key, shared);
      return r ? JSON.parse(r.value) : null;
    } catch {
      return null;
    }
  };
  const set = async (key, value, shared = false) => {
    try {
      await window.storage.set(key, JSON.stringify(value), shared);
    } catch {}
  };
  const del = async (key, shared = false) => {
    try {
      await window.storage.delete(key, shared);
    } catch {}
  };
  return { ready, get, set, del };
}

// Lightweight obfuscation only — this is a client-side demo login, not secure auth.
function encode(str) {
  try {
    return btoa(unescape(encodeURIComponent(str)));
  } catch {
    return str;
  }
}

function formatCountdown(ms) {
  if (ms <= 0) return "00:00:00";
  const h = Math.floor(ms / 3600000);
  const m = Math.floor((ms % 3600000) / 60000);
  const s = Math.floor((ms % 60000) / 1000);
  return [h, m, s].map((n) => String(n).padStart(2, "0")).join(":");
}

// --- Daily quiz engine -----------------------------------------------
// A new quiz is generated for every calendar day, deterministically, from
// a shared question pool. Nothing is ever deleted — each past date keeps
// producing the exact same quiz it always has, so history never changes.

const DAILY_START = "2026-08-15";
const DAILY_COUNT = 60;
// Per-category question counts that sum to DAILY_COUNT (60)
const DAILY_PER_CATEGORY = {
  aptitude: 10,
  verbal: 10,
  technical: 10,
  technical_c: 15,
  technical_java: 15,
};
const DAILY_CATEGORY_KEYS = ["aptitude", "verbal", "technical", "technical_c", "technical_java"];

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function listDailyDates() {
  const dates = [];
  const start = new Date(DAILY_START + "T00:00:00");
  const end = new Date(todayStr() + "T00:00:00");
  for (let d = new Date(end); d >= start; d.setDate(d.getDate() - 1)) {
    const s = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    dates.push(s);
  }
  return dates;
}

function hashStr(str) {
  let h = 2166136261;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

function mulberry32(seed) {
  let a = seed;
  return function () {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function getDailyItems(dateStr, username = "") {
  // Compute how many days since DAILY_START so each date gets a unique offset window.
  const startMs = new Date(DAILY_START + "T00:00:00").getTime();
  const dateMs = new Date(dateStr + "T00:00:00").getTime();
  const dayIndex = Math.max(0, Math.round((dateMs - startMs) / 86400000));

  const picked = [];

  for (const key of DAILY_CATEGORY_KEYS) {
    const cat = QUESTIONS[key];
    const count = DAILY_PER_CATEGORY[key] ?? 10;
    const pool = [...cat.items];

    // Use a larger rotation step to minimize repetition
    // Instead of rotating by count each day, rotate by a larger step
    // This creates a bigger gap between repeated questions
    const rotationStep = Math.max(count, Math.floor(pool.length / 7)); // Ensure at least 7-day gap
    const offset = (dayIndex * rotationStep) % pool.length;
    
    const slice = [];
    for (let i = 0; i < count; i++) {
      slice.push(pool[(offset + i) % pool.length]);
    }

    picked.push(...slice.map((item) => ({ ...item, srcLabel: cat.label })));
  }

  // Final shuffle to mix all categories together — different order each day.
  const mixRng = mulberry32(hashStr(dateStr + "|mix"));
  for (let i = picked.length - 1; i > 0; i--) {
    const j = Math.floor(mixRng() * (i + 1));
    [picked[i], picked[j]] = [picked[j], picked[i]];
  }

  return picked.slice(0, DAILY_COUNT);
}

function getCategoryMeta(catKey, username = "") {
  if (catKey && catKey.startsWith("daily_")) {
    const date = catKey.slice(6);
    return { label: "Daily Quiz", code: date, items: getDailyItems(date, username) };
  }
  return QUESTIONS[catKey];
}

export default function QuizApp() {
  const { ready, get, set, del } = useStorage();
  const [authChecked, setAuthChecked] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [authStage, setAuthStage] = useState("login"); // login | register
  const [authError, setAuthError] = useState("");
  const [authBusy, setAuthBusy] = useState(false);

  const [view, setView] = useState("home"); // home | quiz | result | key | answersMenu | profile | leaderboard
  const [category, setCategory] = useState(null);
  const [activeItems, setActiveItems] = useState([]);
  const [qIndex, setQIndex] = useState(0);
  const [answers, setAnswers] = useState({});
  const [submissions, setSubmissions] = useState({});
  const [leaderboard, setLeaderboard] = useState([]);
  const [onlineUsers, setOnlineUsers] = useState({});
  const [now, setNow] = useState(Date.now());
  const tickRef = useRef(null);
  const presenceRef = useRef(null);

  // Restore session (device-level) and re-attach to the account on load.
  useEffect(() => {
    if (!ready) return;
    (async () => {
      const session = await get("session");
      if (session && session.username) {
        const account = await get(`user:${session.username}`, true);
        if (account) {
          setCurrentUser(session.username);
        }
      }
      setAuthChecked(true);
    })();
  }, [ready]);

  // Load this user's submissions once logged in.
  useEffect(() => {
    if (!currentUser) return;
    (async () => {
      const subs = {};
      for (const key of Object.keys(QUESTIONS)) {
        subs[key] = await get(`submission_${currentUser}_${key}`, true);
      }
      for (const date of listDailyDates()) {
        subs[`daily_${date}`] = await get(
          `submission_${currentUser}_daily_${date}`,
          true,
        );
      }
      setSubmissions(subs);
    })();
  }, [currentUser]);

  useEffect(() => {
    tickRef.current = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(tickRef.current);
  }, []);

  // Shortcut: Ctrl/Cmd + L toggles a logo overlay
  const [logoVisible, setLogoVisible] = useState(false);
  useEffect(() => {
    const onKey = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "l") {
        e.preventDefault();
        setLogoVisible((v) => !v);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  const register = async (username, password, displayName) => {
    setAuthError("");
    const uname = username.trim();
    if (!uname || !password) {
      setAuthError("Enter a username and password.");
      return;
    }
    if (password.length < 4) {
      setAuthError("Password must be at least 4 characters.");
      return;
    }
    setAuthBusy(true);
    const existing = await get(`user:${uname}`, true);
    if (existing) {
      setAuthError("That username is already taken.");
      setAuthBusy(false);
      return;
    }
    const account = {
      password: encode(password),
      createdAt: Date.now(),
      displayName: (displayName || "").trim() || uname,
      provider: 'local'
    };
    await set(`user:${uname}`, account, true);
    
    // Add to all_users list
    const allUsers = await get("all_users", true) || [];
    if (!allUsers.includes(uname)) {
      allUsers.push(uname);
      await set("all_users", allUsers, true);
    }
    
    await set("session", { username: uname }, false);
    setAuthBusy(false);
    setCurrentUser(uname);
    // mark presence for this device/account
    try {
      await set(`presence:${uname}`, { lastSeen: Date.now() }, true);
    } catch {}
    setView("home");
  };

  const login = async (username, password) => {
    setAuthError("");
    const uname = username.trim();
    if (!uname || !password) {
      setAuthError("Enter your username and password.");
      return;
    }
    setAuthBusy(true);
    const account = await get(`user:${uname}`, true);
    if (!account || account.password !== encode(password)) {
      setAuthError("Incorrect username or password.");
      setAuthBusy(false);
      return;
    }
    await set("session", { username: uname }, false);
    setAuthBusy(false);
    setCurrentUser(uname);
    // mark presence for this device/account
    try {
      await set(`presence:${uname}`, { lastSeen: Date.now() }, true);
    } catch {}
    setView("home");
  };

  const logout = async () => {
    // clear session and presence for this device/account
    try {
      if (currentUser) await del(`presence:${currentUser}`, true);
    } catch {}
    await del("session", false);
    setCurrentUser(null);
    setSubmissions({});
    setView("home");
    setAuthStage("login");
  };

  const handleGoogleLogin = async () => {
    setAuthError("");
    setAuthBusy(true);
    
    // Simulate Google OAuth login (demo implementation)
    // In production, this would use actual Google OAuth with proper credentials
    try {
      // Simulate API delay
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // Generate a mock Google user
      const googleUsername = `google_user_${Date.now()}`;
      
      // Store the user account
      const account = {
        username: googleUsername,
        createdAt: Date.now(),
        provider: 'google'
      };
      await set(`user:${googleUsername}`, account, true);
      
      // Add to all_users list
      const allUsers = await get("all_users", true) || [];
      if (!allUsers.includes(googleUsername)) {
        allUsers.push(googleUsername);
        await set("all_users", allUsers, true);
      }
      
      // Set session
      await set("session", { username: googleUsername }, false);
      setCurrentUser(googleUsername);
      
      // Mark presence
      try {
        await set(`presence:${googleUsername}`, { lastSeen: Date.now() }, true);
      } catch {}
      
      setAuthBusy(false);
      setView("home");
    } catch (err) {
      setAuthError("Google login failed. Please try again.");
      setAuthBusy(false);
    }
  };

  const exportUserDataToExcel = async () => {
    try {
      // Collect all user data from storage
      const userData = [];
      
      // Get all user accounts
      const allUsers = await get("all_users", true) || [];
      
      for (const username of allUsers) {
        const account = await get(`user:${username}`, true);
        if (account) {
          userData.push({
            'Username': username,
            'Created At': new Date(account.createdAt).toLocaleString(),
            'Provider': account.provider || 'local',
            'Last Seen': account.lastSeen ? new Date(account.lastSeen).toLocaleString() : 'N/A'
          });
        }
      }
      
      // Create workbook
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(userData);
      
      // Set column widths
      ws['!cols'] = [
        { wch: 20 }, // Username
        { wch: 25 }, // Created At
        { wch: 15 }, // Provider
        { wch: 25 }  // Last Seen
      ];
      
      XLSX.utils.book_append_sheet(wb, ws, "User Data");
      
      // Generate filename with timestamp
      const fileName = `user_data_${new Date().toISOString().slice(0,10)}.xlsx`;
      
      // Download file
      XLSX.writeFile(wb, fileName);
      
    } catch (err) {
      console.error("Error exporting user data:", err);
      alert("Failed to export user data. Please try again.");
    }
  };

  const startQuiz = (cat) => {
    setCategory(cat);
    setAnswers({});
    setQIndex(0);
    const items =
      cat === "technical"
        ? buildTechnicalItems(Date.now())
        : getCategoryMeta(cat, currentUser).items;
    setActiveItems(items);
    setView("quiz");
  };

  const selectOption = (idx) => {
    setAnswers((prev) => ({ ...prev, [qIndex]: idx }));
  };

  const nextQuestion = () => {
    if (qIndex < activeItems.length - 1) {
      setQIndex(qIndex + 1);
    } else {
      submitQuiz();
    }
  };

  const prevQuestion = () => setQIndex((i) => Math.max(0, i - 1));

  const submitQuiz = async () => {
    const items = activeItems;
    let score = 0;
    items.forEach((item, i) => {
      if (answers[i] === item.correct) score += 1;
    });
    const record = {
      answers,
      score,
      total: items.length,
      submittedAt: Date.now(),
      items,
    };
    await set(`submission_${currentUser}_${category}`, record, true);
    setSubmissions((prev) => ({ ...prev, [category]: record }));
    setView("key");
    refreshLeaderboardEntry(currentUser);
  };

  // Recomputes this user's overall stats (across every round + daily quiz
  // they've ever submitted) and writes it to the shared leaderboard.
  const refreshLeaderboardEntry = async (username) => {
    if (!username) return;
    const keys = [
      ...Object.keys(QUESTIONS),
      ...listDailyDates().map((d) => `daily_${d}`),
    ];
    let totalCorrect = 0;
    let totalPossible = 0;
    let attempts = 0;
    for (const k of keys) {
      const rec = await get(`submission_${username}_${k}`, true);
      if (rec) {
        totalCorrect += rec.score;
        totalPossible += rec.total;
        attempts += 1;
      }
    }
    const pct = totalPossible > 0 ? totalCorrect / totalPossible : 0;
    await set(
      `leaderboard:${username}`,
      { totalCorrect, totalPossible, attempts, pct, updatedAt: Date.now() },
      true,
    );
    loadLeaderboard();
  };

  const loadPresence = async () => {
    try {
      const listing = await window.storage.list("presence:", true);
      if (!listing || !listing.keys) {
        setOnlineUsers({});
        return;
      }
      const m = {};
      for (const key of listing.keys) {
        try {
          const r = await get(key, true);
          if (r) m[key.slice("presence:".length)] = r.lastSeen || r;
        } catch {}
      }
      setOnlineUsers(m);
    } catch {}
  };

  const loadLeaderboard = async () => {
    try {
      const listing = await window.storage.list("leaderboard:", true);
      if (!listing || !listing.keys) return;
      const rows = [];
      for (const key of listing.keys) {
        const r = await get(key, true);
        if (r) {
          const username = key.slice("leaderboard:".length);
          const account = await get(`user:${username}`, true);
          rows.push({
            username,
            displayName: getMemberDisplayName(account, username),
            ...r,
          });
        }
      }
      rows.sort((a, b) => b.pct - a.pct || b.totalCorrect - a.totalCorrect);
      setLeaderboard(rows);
      // refresh presence map when loading leaderboard
      loadPresence();
    } catch {}
  };

  useEffect(() => {
    // while viewing leaderboard, poll presence list so UI stays fresh
    if (view === "leaderboard") {
      loadPresence();
      const id = setInterval(loadPresence, 15000);
      return () => clearInterval(id);
    }
  }, [view]);

  useEffect(() => {
    if (view === "leaderboard" || view === "profile" || view === "result")
      loadLeaderboard();
  }, [view]);

  // presence heartbeat for the signed-in user
  useEffect(() => {
    if (!currentUser) {
      if (presenceRef.current) {
        clearInterval(presenceRef.current);
        presenceRef.current = null;
      }
      return;
    }

    const beat = async () => {
      try {
        await set(`presence:${currentUser}`, { lastSeen: Date.now() }, true);
      } catch {}
    };

    // initial mark and periodic heartbeat
    beat();
    presenceRef.current = setInterval(beat, 30000);

    const onUnload = async () => {
      try {
        await del(`presence:${currentUser}`, true);
      } catch {}
    };
    window.addEventListener("beforeunload", onUnload);

    return () => {
      if (presenceRef.current) {
        clearInterval(presenceRef.current);
        presenceRef.current = null;
      }
      window.removeEventListener("beforeunload", onUnload);
      // clear presence when this component unmounts / user switches
      (async () => {
        try {
          await del(`presence:${currentUser}`, true);
        } catch {}
      })();
    };
  }, [currentUser]);

  if (!authChecked) {
    return (
      <div
        style={{
          minHeight: "100vh",
          background: COLORS.ink,
          color: COLORS.paperDim,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontFamily: "'Inter', sans-serif",
        }}
      >
        <style>{FONT_IMPORT}</style>
        Loading…
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div
        style={{
          minHeight: "100vh",
          background: COLORS.ink,
          color: COLORS.paper,
          fontFamily: "'Inter', sans-serif",
        }}
      >
        <style>{FONT_IMPORT}</style>
        {logoVisible && (
          <div
            onClick={() => setLogoVisible(false)}
            style={{
              position: "fixed",
              inset: 0,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              background: "rgba(0,0,0,0.6)",
              zIndex: 9999,
            }}
          >
            <div
              style={{
                background: COLORS.paper,
                padding: 20,
                borderRadius: 12,
                boxShadow: "0 10px 30px rgba(0,0,0,0.5)",
              }}
            >
              <img
                src={LOGO_DATA_URI}
                alt="EDU TECH"
                style={{ width: 160, height: "auto", display: "block" }}
              />
            </div>
          </div>
        )}
        <div style={{ maxWidth: 420, margin: "0 auto", padding: "48px 20px" }}>
          <div style={{ textAlign: "center", marginBottom: 30 }}>
            <div
              style={{
                width: 84,
                height: 84,
                margin: "0 auto 14px",
                background: COLORS.paper,
                borderRadius: 18,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                boxShadow: "0 4px 14px rgba(0,0,0,0.35)",
              }}
            >
              <img
                src={LOGO_DATA_URI}
                alt="EDU TECH"
                style={{ width: 66, height: "auto", display: "block" }}
              />
            </div>
            <div
              style={{
                fontFamily: "'Spectral', serif",
                fontSize: 24,
                fontWeight: 700,
                letterSpacing: 0.5,
              }}
            >
              EDU TECH
            </div>
            <div
              style={{
                fontSize: 11,
                letterSpacing: 2,
                color: COLORS.paperDim,
                textTransform: "uppercase",
                marginTop: 4,
              }}
            >
              Mock Examination Board
            </div>
          </div>
          {authStage === "login" ? (
            <AuthForm
              mode="login"
              busy={authBusy}
              error={authError}
              onSubmit={login}
              onGoogleLogin={handleGoogleLogin}
              onSwitch={() => {
                setAuthStage("register");
                setAuthError("");
              }}
            />
          ) : (
            <AuthForm
              mode="register"
              busy={authBusy}
              error={authError}
              onSubmit={register}
              onSwitch={() => {
                setAuthStage("login");
                setAuthError("");
              }}
            />
          )}
        </div>
      </div>
    );
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        background: COLORS.ink,
        color: COLORS.paper,
        fontFamily: "'Inter', sans-serif",
        display: "flex",
      }}
    >
      <style>{FONT_IMPORT}</style>
      {logoVisible && (
        <div
          onClick={() => setLogoVisible(false)}
          style={{
            position: "fixed",
            inset: 0,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            background: "rgba(0,0,0,0.6)",
            zIndex: 9999,
          }}
        >
          <div
            style={{
              background: COLORS.paper,
              padding: 20,
              borderRadius: 12,
              boxShadow: "0 10px 30px rgba(0,0,0,0.5)",
            }}
          >
            <img
              src={LOGO_DATA_URI}
              alt="EDU TECH"
              style={{ width: 160, height: "auto", display: "block" }}
            />
          </div>
        </div>
      )}
      <Sidebar
        view={view}
        category={category}
        onNav={(v) => setView(v)}
        username={currentUser}
        onLogout={logout}
      />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div
          style={{ maxWidth: 680, margin: "0 auto", padding: "32px 24px 64px" }}
        >
          {view === "home" && (
            <Home
              submissions={submissions}
              onStart={startQuiz}
              onViewKey={(cat) => {
                setCategory(cat);
                setView("key");
              }}
            />
          )}

          {view === "daily" && (
            <DailyMenu
              submissions={submissions}
              onStart={(date) => startQuiz(`daily_${date}`)}
              onViewKey={(date) => {
                setCategory(`daily_${date}`);
                setView("key");
              }}
            />
          )}

          {view === "ide" && <LanguageIde />}

          {view === "assignments" && (
            <Assignments
              submissions={submissions}
              onStart={(cat) => startQuiz(cat)}
              onViewKey={(cat) => {
                setCategory(cat);
                setView("key");
              }}
            />
          )}

          {view === "quiz" && category && (
            <Quiz
              items={activeItems}
              meta={getCategoryMeta(category, currentUser)}
              qIndex={qIndex}
              answers={answers}
              onSelect={selectOption}
              onNext={nextQuestion}
              onPrev={prevQuestion}
            />
          )}

          {view === "result" && category && (
            <Result
              category={category}
              record={submissions[category]}
              leaderboard={leaderboard}
              currentUser={currentUser}
              onViewKey={() => setView("key")}
              onHome={() =>
                setView(category.startsWith("daily_") ? "daily" : "home")
              }
              onRetake={() => startQuiz(category)}
            />
          )}

          {view === "answersMenu" && (
            <AnswersMenu
              submissions={submissions}
              now={now}
              onSelect={(cat) => {
                setCategory(cat);
                setView("key");
              }}
              onGoExams={() => setView("home")}
            />
          )}

          {view === "key" && category && (
            <AnswerKey
              category={category}
              record={submissions[category]}
              now={now}
              currentUser={currentUser}
              onRetake={
                category.startsWith("daily_")
                  ? undefined
                  : () => startQuiz(category)
              }
              onHome={() =>
                setView(category.startsWith("daily_") ? "daily" : "answersMenu")
              }
            />
          )}

          {view === "leaderboard" && (
            <Leaderboard
              leaderboard={leaderboard}
              currentUser={currentUser}
              onlineUsers={onlineUsers}
              onGoExams={() => setView("home")}
            />
          )}

          {view === "profile" && (
            <Profile
              username={currentUser}
              submissions={submissions}
              leaderboard={leaderboard}
              onLogout={logout}
              onGoExams={() => setView("home")}
              onViewLeaderboard={() => setView("leaderboard")}
              onExportUserData={exportUserDataToExcel}
            />
          )}
        </div>
      </div>
    </div>
  );
}

function Assignments({ submissions, onStart, onViewKey }) {
  // Collect only the assignment-tagged entries from QUESTIONS
  const assignmentEntries = Object.entries(QUESTIONS).filter(
    ([, cat]) => cat.isAssignment,
  );

  // Check if current time is before 5pm (17:00)
  const now = new Date();
  const currentHour = now.getHours();
  const isBefore5pm = currentHour < 17;

  return (
    <div className="animate-fade-in">
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginBottom: 6,
        }}
      >
        <ClipboardList size={20} color={COLORS.amber} />
        <div
          style={{
            fontFamily: "'Spectral', serif",
            fontSize: 22,
            fontWeight: 700,
          }}
        >
          Assignments
        </div>
      </div>
      <p
        style={{
          color: COLORS.paperDim,
          fontSize: 13,
          marginBottom: 28,
          lineHeight: 1.6,
        }}
      >
        Complete each assignment and submit — the detailed answer key unlocks
        immediately after you submit.
      </p>

      <div style={{ display: "grid", gap: 16 }}>
        {assignmentEntries.map(([key, cat], idx) => {
          const sub = submissions[key];
          const pct = sub ? Math.round((sub.score / sub.total) * 100) : null;
          const staggerClass = `stagger-${Math.min(idx + 1, 8)}`;

          return (
            <div
              key={key}
              className={`et-hover-card animate-slide-up ${staggerClass}`}
              style={{
                background: COLORS.inkSoft,
                border: `1px solid ${sub ? COLORS.green : COLORS.line}`,
                borderRadius: 10,
                padding: 22,
              }}
            >
              {/* Header row */}
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "flex-start",
                  gap: 12,
                  flexWrap: "wrap",
                }}
              >
                <div style={{ flex: 1, minWidth: 200 }}>
                  <div
                    style={{
                      fontSize: 10,
                      letterSpacing: 2,
                      color: COLORS.amber,
                      fontFamily: "'JetBrains Mono', monospace",
                      marginBottom: 4,
                    }}
                  >
                    {cat.code}
                  </div>
                  <div
                    style={{
                      fontFamily: "'Spectral', serif",
                      fontSize: 20,
                      fontWeight: 600,
                      marginBottom: 6,
                    }}
                  >
                    {cat.label}
                  </div>
                  {cat.description && (
                    <div
                      style={{
                        fontSize: 12,
                        color: COLORS.paperDim,
                        lineHeight: 1.6,
                        marginBottom: 4,
                      }}
                    >
                      {cat.description}
                    </div>
                  )}
                  <div
                    style={{
                      fontSize: 12,
                      color: COLORS.paperDim,
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      marginTop: 2,
                    }}
                  >
                    <ClipboardList size={12} />
                    {cat.items.length} questions · 60 sec per question
                  </div>
                </div>

                {/* Score badge when submitted */}
                {sub && (
                  <div
                    style={{
                      textAlign: "right",
                      fontFamily: "'JetBrains Mono', monospace",
                      flexShrink: 0,
                    }}
                  >
                    <div
                      style={{
                        fontSize: 22,
                        fontWeight: 700,
                        color: pct >= 50 ? COLORS.green : COLORS.clay,
                        lineHeight: 1,
                      }}
                    >
                      {sub.score}/{sub.total}
                    </div>
                    <div
                      style={{
                        fontSize: 11,
                        color: COLORS.paperDim,
                        marginTop: 3,
                        letterSpacing: 1,
                      }}
                    >
                      {pct}% · SUBMITTED
                    </div>
                  </div>
                )}
              </div>

              {/* Progress bar when submitted */}
              {sub && (
                <div
                  style={{
                    height: 5,
                    background: COLORS.ink,
                    borderRadius: 3,
                    overflow: "hidden",
                    margin: "14px 0 16px",
                  }}
                >
                  <div
                    style={{
                      height: "100%",
                      width: `${pct}%`,
                      background: pct >= 50 ? COLORS.green : COLORS.clay,
                      borderRadius: 3,
                      transition: "width 0.5s cubic-bezier(0.4, 0, 0.2, 1)",
                    }}
                  />
                </div>
              )}

              {/* Action buttons */}
              <div
                style={{
                  display: "flex",
                  gap: 10,
                  marginTop: sub ? 0 : 18,
                  flexWrap: "wrap",
                }}
              >
                <button
                  onClick={() => onStart(key)}
                  disabled={key === "assignment_c" && isBefore5pm}
                  className="et-btn-interactive"
                  style={{
                    background:
                      key === "assignment_c" && isBefore5pm
                        ? COLORS.line
                        : COLORS.paper,
                    color: COLORS.ink,
                    border: "none",
                    padding: "9px 18px",
                    borderRadius: 5,
                    fontSize: 13,
                    fontWeight: 600,
                    cursor:
                      key === "assignment_c" && isBefore5pm
                        ? "not-allowed"
                        : "pointer",
                    fontFamily: "'Inter', sans-serif",
                    display: "flex",
                    alignItems: "center",
                    gap: 6,
                    opacity: key === "assignment_c" && isBefore5pm ? 0.5 : 1,
                  }}
                >
                  {sub ? (
                    <>
                      <RotateCcw size={13} /> Retake
                    </>
                  ) : (
                    <>
                      <Play size={13} /> Start assignment
                    </>
                  )}
                </button>

                {sub && (
                  <button
                    onClick={() => onViewKey(key)}
                    className="et-btn-interactive"
                    style={{
                      background: "none",
                      color: COLORS.green,
                      border: `1px solid ${COLORS.green}`,
                      padding: "9px 18px",
                      borderRadius: 5,
                      fontSize: 13,
                      cursor: "pointer",
                      fontFamily: "'Inter', sans-serif",
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                    }}
                  >
                    <Unlock size={13} /> Answer key
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── Pyodide loader (singleton) ──────────────────────────────────────────────
// NOTE: wrapper is named initPyodide to avoid shadowing window.loadPyodide
let _pyodidePromise = null;
function initPyodide() {
  if (_pyodidePromise) return _pyodidePromise;
  _pyodidePromise = new Promise((resolve, reject) => {
    if (window._pyodideReady) {
      resolve();
      return;
    }
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/pyodide.js";
    script.onload = resolve;
    script.onerror = () => reject(new Error("Failed to load Pyodide script"));
    document.head.appendChild(script);
  }).then(async () => {
    if (!window._pyodideInstance) {
      // window.loadPyodide is injected by the Pyodide script above
      window._pyodideInstance = await window.loadPyodide({
        indexURL: "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/",
      });
      window._pyodideReady = true;
    }
  });
  return _pyodidePromise;
}

// ─── Wandbox API runner (C) ───────────────────────────────────────────────────
// Wandbox is a free, no-auth-required online compiler service.
// C compiler: gcc-13.2.0-c  Docs: https://wandbox.org/api/compile.json
async function runWithWandbox(code, stdin = "") {
  const body = {
    compiler: "gcc-13.2.0-c",
    code,
    stdin,
    "compiler-option-raw": "-std=c11 -Wall",
    "runtime-option-raw": "",
    save: false,
  };
  const res = await fetch("https://wandbox.org/api/compile.json", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Wandbox error ${res.status}${text ? ": " + text : ""}`);
  }
  return res.json();
  // Response shape: { status, program_output, program_error, compiler_output,
  //                   compiler_error, compiler_message, signal, ... }
}

// ─── Shared IDE shell ─────────────────────────────────────────────────────────
const LANG_META = {
  python: {
    label: "Python",
    ext: "py",
    color: "#4C8FBD",
    defaultCode: `# Python 3 — runs in-browser via Pyodide
print("Hello, EDU TECH!")

def factorial(n):
    if n <= 1:
        return 1
    return n * factorial(n - 1)

for i in range(1, 8):
    print(f"{i}! = {factorial(i)}")
`,
  },
  c: {
    label: "C",
    ext: "c",
    color: "#8B6FD1",
    defaultCode: `#include <stdio.h>

int factorial(int n) {
    if (n <= 1) return 1;
    return n * factorial(n - 1);
}

int main() {
    printf("Hello, EDU TECH!\\n");
    for (int i = 1; i <= 7; i++) {
        printf("%d! = %d\\n", i, factorial(i));
    }
    return 0;
}
`,
  },
};

function LanguageIde() {
  const [lang, setLang] = useState("python");
  const [codes, setCodes] = useState({
    python: LANG_META.python.defaultCode,
    c: LANG_META.c.defaultCode,
  });
  const [stdin, setStdin] = useState("");
  const [showStdin, setShowStdin] = useState(false);
  const [output, setOutput] = useState("");
  const [status, setStatus] = useState("idle"); // idle | loading | running | done | error
  const [statusMsg, setStatusMsg] = useState("");
  const meta = LANG_META[lang];

  const setCode = (v) => setCodes((p) => ({ ...p, [lang]: v }));

  const runPython = async () => {
    setStatus("loading");
    setStatusMsg("Loading Pyodide runtime…");
    setOutput("");
    try {
      await initPyodide();
      const py = window._pyodideInstance;
      setStatus("running");
      setStatusMsg("Running…");

      // Redirect stdout/stderr inside Python, collect into a string
      const captureCode = `
import sys, traceback
from io import StringIO as _SIO
_out = _SIO()
_old_out, _old_err = sys.stdout, sys.stderr
sys.stdout = sys.stderr = _out
try:
    exec(compile(${JSON.stringify(codes.python)}, "<stdin>", "exec"))
except SystemExit:
    pass
except Exception:
    traceback.print_exc()
finally:
    sys.stdout = _old_out
    sys.stderr = _old_err
_out.getvalue()
`;
      const result = await py.runPythonAsync(captureCode);
      setOutput(result ?? "(no output)");
      setStatus("done");
      setStatusMsg("Finished");
    } catch (e) {
      setOutput(String(e));
      setStatus("error");
      setStatusMsg("Error");
    }
  };

  const runC = async () => {
    setStatus("loading");
    setStatusMsg("Sending to compiler…");
    setOutput("");
    try {
      const result = await runWithWandbox(codes.c, stdin);
      // Wandbox response: { status, program_output, program_error,
      //                     compiler_output, compiler_error, compiler_message }
      // status "0" = success, non-zero or "CE" = compile error
      const compileError = result.compiler_error || "";
      const compileOutput = result.compiler_output || "";

      // A compile error means no program ran
      if (
        result.status === "CE" ||
        (compileError && !result.program_output && !result.program_error)
      ) {
        const msg = (
          compileError ||
          compileOutput ||
          "Unknown compile error"
        ).trim();
        setOutput("Compile error:\n" + msg);
        setStatus("error");
        setStatusMsg("Compile error");
        return;
      }

      // Program ran — show stdout + stderr
      const stdout = result.program_output || "";
      const stderr = result.program_error || "";
      // Show compiler warnings if any (non-fatal)
      const warnings = compileError
        ? "Compiler warnings:\n" + compileError + "\n\n"
        : "";
      const combined =
        warnings +
        stdout +
        (stderr ? (stdout ? "\nstderr:\n" : "stderr:\n") + stderr : "");
      const exitCode = parseInt(result.status, 10);
      setOutput(combined || "(no output)");
      setStatus(exitCode === 0 ? "done" : "error");
      setStatusMsg(exitCode === 0 ? "Exited 0" : `Exited ${exitCode}`);
    } catch (e) {
      setOutput(`Could not reach the compiler.\n\nDetails: ${e.message}`);
      setStatus("error");
      setStatusMsg("Network error");
    }
  };

  const handleRun = () => {
    if (status === "loading" || status === "running") return;
    if (lang === "python") runPython();
    else runC();
  };

  const busy = status === "loading" || status === "running";

  const statusColor = {
    idle: COLORS.paperDim,
    loading: COLORS.amber,
    running: COLORS.amber,
    done: COLORS.green,
    error: COLORS.clay,
  }[status];

  return (
    <div className="animate-fade-in" style={{ padding: "8px 0" }}>
      {/* Header */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginBottom: 16,
          flexWrap: "wrap",
        }}
      >
        <Terminal size={20} color={COLORS.amber} />
        <div
          style={{
            fontFamily: "'Spectral', serif",
            fontSize: 22,
            fontWeight: 700,
          }}
        >
          Code IDE
        </div>
        {/* Language tabs */}
        <div style={{ display: "flex", gap: 6, marginLeft: "auto" }}>
          {Object.entries(LANG_META).map(([key, m]) => (
            <button
              key={key}
              onClick={() => {
                setLang(key);
                setOutput("");
                setStatus("idle");
                setStatusMsg("");
              }}
              className="et-btn-interactive"
              style={{
                padding: "6px 14px",
                borderRadius: 6,
                border: `1px solid ${lang === key ? m.color : COLORS.line}`,
                background: lang === key ? `${m.color}22` : "transparent",
                color: lang === key ? m.color : COLORS.paperDim,
                fontWeight: 600,
                fontSize: 13,
                cursor: "pointer",
                fontFamily: "'Inter', sans-serif",
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
            >
              <Code2 size={13} />
              {m.label}
            </button>
          ))}
        </div>
      </div>

      <div
        className="animate-slide-up"
        style={{
          background: COLORS.inkSoft,
          border: `1px solid ${COLORS.line}`,
          borderRadius: 10,
          overflow: "hidden",
        }}
      >
        {/* Toolbar */}
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            padding: "10px 16px",
            borderBottom: `1px solid ${COLORS.line}`,
            background: COLORS.ink,
            flexWrap: "wrap",
          }}
        >
          <div
            style={{
              fontSize: 11,
              letterSpacing: 1.5,
              color: meta.color,
              fontFamily: "'JetBrains Mono', monospace",
              textTransform: "uppercase",
            }}
          >
            {meta.label} · main.{meta.ext}
          </div>
          <div style={{ flex: 1 }} />

          {/* stdin toggle (C only) */}
          {lang === "c" && (
            <button
              onClick={() => setShowStdin((v) => !v)}
              className="et-btn-interactive"
              style={{
                background: showStdin ? `${COLORS.teal}22` : "transparent",
                border: `1px solid ${showStdin ? COLORS.teal : COLORS.line}`,
                color: showStdin ? COLORS.teal : COLORS.paperDim,
                borderRadius: 5,
                padding: "5px 10px",
                fontSize: 11,
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: 5,
                fontFamily: "'Inter', sans-serif",
              }}
            >
              stdin{" "}
              <ChevronDown
                size={11}
                style={{ transform: showStdin ? "rotate(180deg)" : "none", transition: "transform 0.2s ease" }}
              />
            </button>
          )}

          {/* Clear */}
          <button
            onClick={() => {
              setCode("");
              setOutput("");
              setStatus("idle");
              setStatusMsg("");
            }}
            className="et-btn-interactive"
            title="Clear editor"
            style={{
              background: "transparent",
              border: `1px solid ${COLORS.line}`,
              color: COLORS.paperDim,
              borderRadius: 5,
              padding: "5px 8px",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
            }}
          >
            <Trash2 size={13} />
          </button>

          {/* Run */}
          <button
            onClick={handleRun}
            disabled={busy}
            className="et-btn-interactive"
            style={{
              background: busy ? COLORS.line : COLORS.amber,
              color: busy ? COLORS.paperDim : COLORS.ink,
              border: "none",
              borderRadius: 6,
              padding: "7px 16px",
              fontWeight: 700,
              fontSize: 13,
              cursor: busy ? "not-allowed" : "pointer",
              display: "flex",
              alignItems: "center",
              gap: 6,
              fontFamily: "'Inter', sans-serif",
            }}
          >
            {busy ? (
              <Loader2
                size={13}
                style={{ animation: "spin 1s linear infinite" }}
              />
            ) : (
              <Play size={13} />
            )}
            {busy ? statusMsg : "Run"}
          </button>
        </div>

        {/* stdin panel */}
        {lang === "c" && showStdin && (
          <div
            style={{
              padding: "10px 16px",
              borderBottom: `1px solid ${COLORS.line}`,
              background: COLORS.ink,
            }}
          >
            <div
              style={{
                fontSize: 11,
                letterSpacing: 1,
                color: COLORS.teal,
                marginBottom: 6,
                textTransform: "uppercase",
              }}
            >
              stdin input
            </div>
            <textarea
              value={stdin}
              onChange={(e) => setStdin(e.target.value)}
              placeholder="Provide stdin for your program…"
              spellCheck={false}
              style={{
                width: "100%",
                height: 72,
                background: COLORS.inkSoft,
                color: COLORS.paper,
                border: `1px solid ${COLORS.line}`,
                borderRadius: 6,
                padding: "8px 10px",
                fontFamily: "'JetBrains Mono', monospace",
                fontSize: 12,
                resize: "vertical",
                boxSizing: "border-box",
              }}
            />
          </div>
        )}

        {/* Code editor */}
        <textarea
          value={codes[lang]}
          onChange={(e) => setCode(e.target.value)}
          spellCheck={false}
          style={{
            width: "100%",
            minHeight: 320,
            background: "#0d1117",
            color: "#e6edf3",
            border: "none",
            borderBottom: `1px solid ${COLORS.line}`,
            padding: "16px 18px",
            fontFamily: "'JetBrains Mono', monospace",
            fontSize: 13,
            lineHeight: 1.65,
            resize: "vertical",
            boxSizing: "border-box",
            outline: "none",
            tabSize: 4,
          }}
          onKeyDown={(e) => {
            // Tab inserts 4 spaces
            if (e.key === "Tab") {
              e.preventDefault();
              const ta = e.target;
              const start = ta.selectionStart;
              const end = ta.selectionEnd;
              const newVal =
                ta.value.substring(0, start) + "    " + ta.value.substring(end);
              setCode(newVal);
              requestAnimationFrame(() => {
                ta.selectionStart = ta.selectionEnd = start + 4;
              });
            }
          }}
        />

        {/* Output panel */}
        <div style={{ padding: "12px 16px", background: "#0d1117" }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              marginBottom: 8,
            }}
          >
            <Terminal size={12} color={COLORS.paperDim} />
            <span
              style={{
                fontSize: 11,
                letterSpacing: 1.2,
                color: COLORS.paperDim,
                textTransform: "uppercase",
              }}
            >
              Output
            </span>
            {statusMsg && (
              <span
                style={{
                  fontSize: 11,
                  color: statusColor,
                  fontFamily: "'JetBrains Mono', monospace",
                  marginLeft: 6,
                }}
              >
                — {statusMsg}
              </span>
            )}
            {lang === "python" && status === "idle" && (
              <span
                style={{
                  fontSize: 11,
                  color: COLORS.paperDim,
                  marginLeft: "auto",
                  opacity: 0.6,
                }}
              >
                Runs in-browser via Pyodide
              </span>
            )}
            {lang === "c" && status === "idle" && (
              <span
                style={{
                  fontSize: 11,
                  color: COLORS.paperDim,
                  marginLeft: "auto",
                  opacity: 0.6,
                }}
              >
                Compiled via Wandbox (gcc 13)
              </span>
            )}
          </div>
          <pre
            style={{
              margin: 0,
              whiteSpace: "pre-wrap",
              wordBreak: "break-word",
              color: status === "error" ? COLORS.clay : "#e6edf3",
              fontFamily: "'JetBrains Mono', monospace",
              fontSize: 13,
              lineHeight: 1.6,
              minHeight: 80,
            }}
          >
            {busy
              ? "⟳  " + statusMsg + "…"
              : output || "Press Run to execute your code."}
          </pre>
        </div>
      </div>

      {/* CSS spin keyframe injected once */}
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

// Keep PythonIde as alias so nothing else breaks (unused but safe)
function PythonIde() {
  return <LanguageIde />;
}

function Sidebar({ view, category, onNav, username, onLogout }) {
  const [narrow, setNarrow] = useState(
    typeof window !== "undefined" && window.innerWidth < 640,
  );

  useEffect(() => {
    const onResize = () => setNarrow(window.innerWidth < 640);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  const items = [
    { id: "home", label: "Exams", icon: BookOpen },
    { id: "daily", label: "Daily", icon: CalendarDays },
    { id: "assignments", label: "Assignments", icon: ClipboardList },
    { id: "answersMenu", label: "Answers", icon: KeyRound },
    { id: "ide", label: "Code IDE", icon: Code2 },
    { id: "leaderboard", label: "Leaders", icon: Trophy },
    { id: "profile", label: "Profile", icon: User },
  ];

  const isDaily = category && category.startsWith("daily_");
  const isAssignment = category && getCategoryMeta(category, username)?.isAssignment;
  const activeTab =
    view === "quiz" || view === "result"
      ? isDaily
        ? "daily"
        : "home"
      : view === "key"
        ? isDaily
          ? "daily"
          : "answersMenu"
        : view;

  return (
    <div
      style={{
        width: narrow ? 68 : 216,
        flexShrink: 0,
        borderRight: `1px solid ${COLORS.line}`,
        display: "flex",
        flexDirection: "column",
        minHeight: "100vh",
        padding: narrow ? "20px 8px" : "24px 16px",
        boxSizing: "border-box",
        overflow: "hidden",
      }}
    >
      <div
        onClick={() => onNav("home")}
        style={{
          cursor: "pointer",
          marginBottom: 30,
          display: "flex",
          flexDirection: narrow ? "column" : "row",
          alignItems: "center",
          gap: narrow ? 0 : 10,
        }}
      >
        <div
          style={{
            width: narrow ? 36 : 42,
            height: narrow ? 36 : 42,
            flexShrink: 0,
            background: COLORS.paper,
            borderRadius: narrow ? 9 : 10,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <img
            src={LOGO_DATA_URI}
            alt="EDU TECH"
            style={{
              width: narrow ? 27 : 32,
              height: "auto",
              display: "block",
            }}
          />
        </div>
        {!narrow && (
          <div>
            <div
              style={{
                fontFamily: "'Spectral', serif",
                fontSize: 17,
                fontWeight: 700,
                letterSpacing: 0.4,
                lineHeight: 1.1,
              }}
            >
              EDU TECH
            </div>
            <div
              style={{
                fontSize: 9,
                letterSpacing: 1.2,
                color: COLORS.paperDim,
                textTransform: "uppercase",
                marginTop: 3,
              }}
            >
              Mock Exam Board
            </div>
          </div>
        )}
      </div>

      <div
        style={{ display: "flex", flexDirection: "column", gap: 4, flex: 1 }}
      >
        {items.map((it) => {
          const isActive = activeTab === it.id;
          const Icon = it.icon;
          return (
            <button
              key={it.id}
              onClick={() => onNav(it.id)}
              title={it.label}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 12,
                justifyContent: narrow ? "center" : "flex-start",
                background: isActive ? "rgba(217,164,65,0.12)" : "transparent",
                border: "none",
                borderLeft: `2px solid ${isActive ? COLORS.amber : "transparent"}`,
                color: isActive ? COLORS.amber : COLORS.paperDim,
                padding: narrow ? "10px 0" : "10px 12px",
                borderRadius: 6,
                fontSize: 13,
                fontWeight: 600,
                cursor: "pointer",
                fontFamily: "'Inter', sans-serif",
                textAlign: "left",
              }}
            >
              <Icon size={17} style={{ flexShrink: 0 }} />
              {!narrow && it.label}
            </button>
          );
        })}
      </div>

      {username && (
        <div
          style={{
            borderTop: `1px solid ${COLORS.line}`,
            paddingTop: 14,
            marginTop: 14,
          }}
        >
          {!narrow && (
            <div
              onClick={() => onNav("profile")}
              style={{
                fontSize: 12,
                color: COLORS.paperDim,
                cursor: "pointer",
                fontFamily: "'JetBrains Mono', monospace",
                marginBottom: 10,
                wordBreak: "break-all",
              }}
            >
              {username}
            </div>
          )}
          <button
            onClick={onLogout}
            title="Log out"
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              justifyContent: narrow ? "center" : "flex-start",
              width: "100%",
              background: "none",
              border: `1px solid ${COLORS.line}`,
              color: COLORS.paperDim,
              fontSize: 12,
              padding: "8px 10px",
              borderRadius: 5,
              cursor: "pointer",
              fontFamily: "'Inter', sans-serif",
              boxSizing: "border-box",
            }}
          >
            <LogOut size={14} />
            {!narrow && "Log out"}
          </button>
        </div>
      )}
    </div>
  );
}

function AuthForm({ mode, busy, error, onSubmit, onSwitch, onGoogleLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const isRegister = mode === "register";

  const handleSubmit = (e) => {
    e.preventDefault();
    if (isRegister && password !== confirm) {
      return;
    }
    onSubmit(username, password, undefined);
  };

  const handleGoogleLogin = (e) => {
    e.preventDefault();
    if (onGoogleLogin) {
      onGoogleLogin();
    }
  };

  const mismatch = isRegister && confirm.length > 0 && password !== confirm;

  return (
    <div
      style={{
        background: COLORS.inkSoft,
        border: `1px solid ${COLORS.line}`,
        borderRadius: 10,
        padding: 24,
      }}
    >
      <div
        style={{
          fontFamily: "'Spectral', serif",
          fontSize: 19,
          fontWeight: 600,
          marginBottom: 4,
        }}
      >
        {isRegister ? "Create your account" : ""}
      </div>
      {isRegister && (
        <p style={{ color: COLORS.paperDim, fontSize: 13, marginBottom: 20 }}>
          Register to start your rounds and track your scores.
        </p>
      )}

      <div onSubmit={handleSubmit}>
        <FieldLabel text={isRegister ? "Login username" : "Enter Your Name"} />
        <input
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          placeholder={
            isRegister ? "Choose a login username" : "Enter Your Name"
          }
          style={inputStyle}
          autoCapitalize="none"
        />

        <FieldLabel text="Password" />
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          placeholder="At least 4 characters"
          style={inputStyle}
        />

        {isRegister && (
          <>
            <FieldLabel text="Confirm password" />
            <input
              type="password"
              value={confirm}
              onChange={(e) => setConfirm(e.target.value)}
              placeholder="Re-enter password"
              style={{
                ...inputStyle,
                borderColor: mismatch ? COLORS.clay : COLORS.line,
              }}
            />
            {mismatch && (
              <div
                style={{
                  color: COLORS.clay,
                  fontSize: 12,
                  marginTop: -8,
                  marginBottom: 12,
                }}
              >
                Passwords don't match.
              </div>
            )}
          </>
        )}

        {error && (
          <div style={{ color: COLORS.clay, fontSize: 13, marginBottom: 14 }}>
            {error}
          </div>
        )}

        <button
          onClick={handleSubmit}
          disabled={busy || mismatch}
          style={{
            ...btnPrimary,
            width: "100%",
            justifyContent: "center",
            opacity: busy ? 0.7 : 1,
            marginTop: 4,
          }}
        >
          {busy ? "Please wait…" : isRegister ? "Register" : "Log in"}
        </button>

        {!isRegister && (
          <>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                margin: "20px 0",
                color: COLORS.paperDim,
              }}
            >
              <div style={{ flex: 1, height: "1px", background: COLORS.line }} />
              <div style={{ padding: "0 12px", fontSize: 12 }}>OR</div>
              <div style={{ flex: 1, height: "1px", background: COLORS.line }} />
            </div>

            <button
              onClick={handleGoogleLogin}
              disabled={busy}
              style={{
                width: "100%",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 10,
                background: "#fff",
                color: "#333",
                border: `1px solid ${COLORS.line}`,
                borderRadius: 6,
                padding: "10px 12px",
                fontSize: 14,
                fontFamily: "'Inter', sans-serif",
                cursor: busy ? "not-allowed" : "pointer",
                opacity: busy ? 0.7 : 1,
                outline: "none",
              }}
            >
              <svg width="18" height="18" viewBox="0 0 18 18">
                <path
                  d="M17.64 9.2c0-.637-.057-1.251-.164-1.84H9v3.481h4.844a4.14 4.14 0 0 1-1.796 2.716v2.258h2.908c1.702-1.567 2.684-3.874 2.684-6.615z"
                  fill="#4285F4"
                />
                <path
                  d="M9 18c2.43 0 4.467-.806 5.956-2.18l-2.908-2.259c-.806.54-1.836.86-3.048.86-2.344 0-4.328-1.584-5.036-3.715H.957v2.332A8.997 8.997 0 0 0 9 18z"
                  fill="#34A853"
                />
                <path
                  d="M3.964 10.71A5.41 5.41 0 0 1 3.682 9c0-.593.102-1.17.282-1.71V4.958H.957A8.996 8.996 0 0 0 0 9c0 1.452.348 2.827.957 4.042l3.007-2.332z"
                  fill="#FBBC05"
                />
                <path
                  d="M9 3.58c1.321 0 2.508.454 3.44 1.345l2.582-2.58C13.463.891 11.426 0 9 0A8.997 8.997 0 0 0 .957 4.958L3.964 7.29C4.672 5.159 6.656 3.58 9 3.58z"
                  fill="#EA4335"
                />
              </svg>
              Continue with Google
            </button>
          </>
        )}
      </div>

      <div
        style={{
          textAlign: "center",
          marginTop: 18,
          fontSize: 13,
          color: COLORS.paperDim,
        }}
      >
        {isRegister ? "Already have an account?" : "New here?"}{" "}
        <span
          onClick={onSwitch}
          style={{ color: COLORS.amber, cursor: "pointer", fontWeight: 600 }}
        >
          {isRegister ? "Log in" : "Register"}
        </span>
      </div>
    </div>
  );
}

function FieldLabel({ text }) {
  return (
    <div
      style={{
        fontSize: 11,
        letterSpacing: 1,
        color: COLORS.paperDim,
        textTransform: "uppercase",
        marginBottom: 6,
      }}
    >
      {text}
    </div>
  );
}

const inputStyle = {
  width: "100%",
  boxSizing: "border-box",
  background: COLORS.ink,
  border: `1px solid ${COLORS.line}`,
  borderRadius: 6,
  padding: "10px 12px",
  color: COLORS.paper,
  fontSize: 14,
  fontFamily: "'Inter', sans-serif",
  marginBottom: 16,
  outline: "none",
};

function Profile({
  username,
  submissions,
  leaderboard,
  onLogout,
  onGoExams,
  onViewLeaderboard,
  onExportUserData,
}) {
  const attempted = Object.entries(submissions).filter(([, v]) => v);
  const totalScore = attempted.reduce((sum, [, v]) => sum + v.score, 0);
  const totalPossible = attempted.reduce((sum, [, v]) => sum + v.total, 0);
  const totalWrong = attempted.reduce((sum, [, v]) => {
    const answeredCount = Object.keys(v.answers || {}).length;
    return sum + (answeredCount - v.score);
  }, 0);
  const totalUnanswered = attempted.reduce((sum, [, v]) => {
    const answeredCount = Object.keys(v.answers || {}).length;
    return sum + (v.total - answeredCount);
  }, 0);
  const overallPct =
    totalPossible > 0 ? Math.round((totalScore / totalPossible) * 100) : 0;

  const myRank =
    leaderboard && leaderboard.length
      ? leaderboard.findIndex((r) => r.username === username) + 1
      : 0;

  const statusFor = (pct) => {
    if (pct >= 75) return { label: "Strong", color: COLORS.green };
    if (pct >= 50) return { label: "Average", color: COLORS.amber };
    return { label: "Needs work", color: COLORS.clay };
  };

  return (
    <div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 14,
          marginBottom: 24,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            width: 56,
            height: 56,
            borderRadius: "50%",
            background: COLORS.amber,
            color: COLORS.ink,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontFamily: "'Spectral', serif",
            fontSize: 24,
            fontWeight: 700,
            flexShrink: 0,
          }}
        >
          {username.charAt(0).toUpperCase()}
        </div>
        <div style={{ flex: 1, minWidth: 140 }}>
          <div
            style={{
              fontFamily: "'Spectral', serif",
              fontSize: 20,
              fontWeight: 600,
            }}
          >
            {username}
          </div>
          <div style={{ fontSize: 12, color: COLORS.paperDim, marginTop: 2 }}>
            {attempted.length} of {Object.keys(QUESTIONS).length} rounds
            completed
          </div>
        </div>
        <button
          onClick={onViewLeaderboard}
          style={{ ...btnGhost, display: "flex", alignItems: "center", gap: 7 }}
        >
          <Trophy size={14} color={COLORS.amber} />
          {myRank > 0 ? `Rank #${myRank} overall` : "View leaderboard"}
        </button>
      </div>

      <div
        style={{
          fontSize: 11,
          letterSpacing: 2,
          color: COLORS.paperDim,
          textTransform: "uppercase",
          marginBottom: 10,
        }}
      >
        Overall analysis
      </div>

      {totalPossible > 0 ? (
        <>
          <div
            style={{
              background: COLORS.inkSoft,
              border: `1px solid ${COLORS.line}`,
              borderRadius: 8,
              padding: 20,
              marginBottom: 14,
              display: "flex",
              alignItems: "center",
              gap: 16,
              flexWrap: "wrap",
            }}
          >
            <div
              style={{
                position: "relative",
                width: 168,
                height: 168,
                flexShrink: 0,
              }}
            >
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={[
                      ...Object.entries(QUESTIONS)
                        .filter(([key]) => submissions[key])
                        .map(([key, cat]) => ({
                          name: cat.label,
                          value: submissions[key].score,
                          color: CATEGORY_COLOR[key],
                        })),
                      {
                        name: "Remaining",
                        value: Math.max(totalPossible - totalScore, 0),
                        color: COLORS.line,
                      },
                    ]}
                    dataKey="value"
                    innerRadius={58}
                    outerRadius={78}
                    startAngle={90}
                    endAngle={-270}
                    paddingAngle={totalScore > 0 ? 3 : 0}
                    stroke="none"
                  >
                    {Object.entries(QUESTIONS)
                      .filter(([key]) => submissions[key])
                      .map(([key]) => (
                        <Cell key={key} fill={CATEGORY_COLOR[key]} />
                      ))}
                    <Cell key="remaining" fill={COLORS.line} />
                  </Pie>
                </PieChart>
              </ResponsiveContainer>
              <div
                style={{
                  position: "absolute",
                  top: "50%",
                  left: "50%",
                  transform: "translate(-50%, -50%)",
                  textAlign: "center",
                  pointerEvents: "none",
                }}
              >
                <div
                  style={{
                    fontFamily: "'JetBrains Mono', monospace",
                    fontSize: 26,
                    fontWeight: 700,
                    color: COLORS.paper,
                    lineHeight: 1,
                  }}
                >
                  {overallPct}%
                </div>
                <div
                  style={{
                    fontSize: 10,
                    color: COLORS.paperDim,
                    marginTop: 4,
                    letterSpacing: 0.5,
                  }}
                >
                  {totalScore}/{totalPossible} solved
                </div>
              </div>
            </div>

            <div style={{ flex: 1, minWidth: 140 }}>
              <div
                style={{
                  display: "inline-block",
                  fontSize: 11,
                  fontWeight: 600,
                  letterSpacing: 1,
                  color: statusFor(overallPct).color,
                  border: `1px solid ${statusFor(overallPct).color}`,
                  borderRadius: 20,
                  padding: "3px 10px",
                  marginBottom: 12,
                }}
              >
                {statusFor(overallPct).label.toUpperCase()}
              </div>
              <div style={{ display: "grid", gap: 8 }}>
                {Object.entries(QUESTIONS).map(([key, cat]) => {
                  const sub = submissions[key];
                  return (
                    <div
                      key={key}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        fontSize: 12,
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 7,
                        }}
                      >
                        <span
                          style={{
                            width: 8,
                            height: 8,
                            borderRadius: "50%",
                            background: CATEGORY_COLOR[key],
                            display: "inline-block",
                          }}
                        />
                        <span style={{ color: COLORS.paperDim }}>
                          {cat.label}
                        </span>
                      </div>
                      <span
                        style={{
                          fontFamily: "'JetBrains Mono', monospace",
                          color: sub ? COLORS.paper : COLORS.paperDim,
                        }}
                      >
                        {sub ? `${sub.score}/${sub.total}` : "—"}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr 1fr",
              gap: 10,
              marginBottom: 24,
            }}
          >
            <StatChip label="Correct" value={totalScore} color={COLORS.green} />
            <StatChip
              label="Incorrect"
              value={totalWrong}
              color={COLORS.clay}
            />
            <StatChip
              label="Unanswered"
              value={totalUnanswered}
              color={COLORS.paperDim}
            />
          </div>
        </>
      ) : (
        <div
          style={{
            textAlign: "center",
            padding: "24px 16px",
            border: `1px dashed ${COLORS.line}`,
            borderRadius: 8,
            color: COLORS.paperDim,
            fontSize: 13,
            marginBottom: 24,
          }}
        >
          Complete a round to see your analysis here.
        </div>
      )}

      <div
        style={{
          fontSize: 11,
          letterSpacing: 2,
          color: COLORS.paperDim,
          textTransform: "uppercase",
          marginBottom: 10,
        }}
      >
        Round-wise performance
      </div>
      <div style={{ display: "grid", gap: 12, marginBottom: 24 }}>
        {Object.entries(QUESTIONS).map(([key, cat]) => {
          const sub = submissions[key];
          const pct = sub ? Math.round((sub.score / sub.total) * 100) : null;
          return (
            <div
              key={key}
              style={{
                background: COLORS.inkSoft,
                border: `1px solid ${COLORS.line}`,
                borderRadius: 8,
                padding: 16,
              }}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <div>
                  <div
                    style={{
                      fontSize: 10,
                      letterSpacing: 2,
                      color: COLORS.amber,
                      fontFamily: "'JetBrains Mono', monospace",
                    }}
                  >
                    {cat.code}
                  </div>
                  <div
                    style={{
                      fontFamily: "'Spectral', serif",
                      fontSize: 16,
                      marginTop: 2,
                    }}
                  >
                    {cat.label}
                  </div>
                </div>
                {sub ? (
                  <div style={{ textAlign: "right" }}>
                    <div
                      style={{
                        fontFamily: "'JetBrains Mono', monospace",
                        fontSize: 16,
                        fontWeight: 700,
                        color: statusFor(pct).color,
                      }}
                    >
                      {sub.score}/{sub.total}
                    </div>
                    <div
                      style={{
                        fontSize: 11,
                        color: COLORS.paperDim,
                        marginTop: 2,
                      }}
                    >
                      {pct}%
                    </div>
                  </div>
                ) : (
                  <div style={{ fontSize: 12, color: COLORS.paperDim }}>
                    Not attempted
                  </div>
                )}
              </div>
              {sub && (
                <div
                  style={{
                    height: 6,
                    background: COLORS.ink,
                    borderRadius: 3,
                    overflow: "hidden",
                    marginTop: 12,
                  }}
                >
                  <div
                    style={{
                      height: "100%",
                      width: `${pct}%`,
                      background: statusFor(pct).color,
                      borderRadius: 3,
                      transition: "width 0.4s ease",
                    }}
                  />
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div
        style={{
          display: "flex",
          gap: 10,
          justifyContent: "center",
          flexWrap: "wrap",
        }}
      >
        <button onClick={onGoExams} style={btnPrimary}>
          Go to exams
        </button>
        {onExportUserData && (
          <button onClick={onExportUserData} style={btnGhost}>
            <Download size={14} style={{ marginRight: 6 }} />
            Export User Data
          </button>
        )}
        <button onClick={onLogout} style={btnGhost}>
          Log out
        </button>
      </div>
    </div>
  );
}

function StatChip({ label, value, color }) {
  return (
    <div
      style={{
        background: COLORS.inkSoft,
        border: `1px solid ${COLORS.line}`,
        borderRadius: 8,
        padding: "12px 8px",
        textAlign: "center",
      }}
    >
      <div
        style={{
          fontFamily: "'JetBrains Mono', monospace",
          fontSize: 18,
          fontWeight: 700,
          color,
        }}
      >
        {value}
      </div>
      <div
        style={{
          fontSize: 10,
          color: COLORS.paperDim,
          marginTop: 3,
          letterSpacing: 0.5,
        }}
      >
        {label}
      </div>
    </div>
  );
}

const RANK_MEDAL = { 1: "#E8C34A", 2: "#C7CDD6", 3: "#C98A4B" };

function Leaderboard({
  leaderboard,
  currentUser,
  onlineUsers = {},
  onGoExams,
}) {
  return (
    <div className="animate-fade-in">
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginBottom: 6,
        }}
      >
        <Trophy size={20} color={COLORS.amber} />
        <div
          style={{
            fontFamily: "'Spectral', serif",
            fontSize: 22,
            fontWeight: 700,
          }}
        >
          Overall leaderboard
        </div>
      </div>
      <p
        style={{
          color: COLORS.paperDim,
          fontSize: 13,
          marginBottom: 22,
          lineHeight: 1.6,
        }}
      >
        Ranked by overall percentage across every round and daily quiz each
        member has submitted.
      </p>

      {leaderboard.length === 0 ? (
        <div
          className="animate-scale-in"
          style={{
            textAlign: "center",
            padding: "40px 20px",
            border: `1px dashed ${COLORS.line}`,
            borderRadius: 8,
            color: COLORS.paperDim,
          }}
        >
          No submissions yet — be the first to appear here.
        </div>
      ) : (
        <div style={{ display: "grid", gap: 8 }}>
          {/* Currently online users summary */}
          {Object.keys(onlineUsers).length > 0 && (
            <div
              className="animate-fade-in"
              style={{
                display: "flex",
                gap: 8,
                alignItems: "center",
                marginBottom: 6,
                flexWrap: "wrap",
              }}
            >
              <div style={{ fontSize: 12, color: COLORS.paperDim }}>
                Online:
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {Object.keys(onlineUsers).map((u) => {
                  const last = onlineUsers[u];
                  const isNow = Date.now() - (last || 0) < 90 * 1000;
                  return (
                    <div
                      key={u}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        padding: "4px 8px",
                        background: isNow
                          ? "rgba(63,143,99,0.08)"
                          : "transparent",
                        border: `1px solid ${isNow ? COLORS.green : COLORS.line}`,
                        borderRadius: 999,
                        fontSize: 12,
                        color: isNow ? COLORS.green : COLORS.paperDim,
                      }}
                    >
                      <div
                        className={isNow ? "animate-online-pulse" : ""}
                        style={{
                          width: 8,
                          height: 8,
                          borderRadius: "50%",
                          background: isNow ? COLORS.green : COLORS.paperDim,
                          flexShrink: 0,
                        }}
                      />
                      <div
                        style={{ fontFamily: "'JetBrains Mono', monospace" }}
                      >
                        {u}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
          {leaderboard.map((row, i) => {
            const rank = i + 1;
            const isMe = row.username === currentUser;
            const pct = Math.round(row.pct * 100);
            const staggerClass = `stagger-${Math.min(i + 1, 8)}`;
            return (
              <div
                key={row.username}
                className={`et-hover-row animate-slide-up ${staggerClass}`}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 14,
                  background: isMe ? "rgba(217,164,65,0.10)" : COLORS.inkSoft,
                  border: `1px solid ${isMe ? COLORS.amber : COLORS.line}`,
                  borderRadius: 8,
                  padding: "12px 16px",
                }}
              >
                <div
                  style={{
                    width: 28,
                    textAlign: "center",
                    fontFamily: "'JetBrains Mono', monospace",
                    fontSize: 14,
                    fontWeight: 700,
                    color: RANK_MEDAL[rank] || COLORS.paperDim,
                    flexShrink: 0,
                  }}
                >
                  {rank}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div
                    style={{
                      fontSize: 14,
                      fontWeight: isMe ? 700 : 500,
                      fontFamily: "'JetBrains Mono', monospace",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {row.displayName || row.username}
                    {isMe && " (you)"}
                  </div>
                  <div
                    style={{
                      fontSize: 11,
                      color: COLORS.paperDim,
                      marginTop: 2,
                    }}
                  >
                    {row.totalCorrect}/{row.totalPossible} correct ·{" "}
                    {row.attempts} attempt{row.attempts === 1 ? "" : "s"}
                  </div>
                </div>
                <div
                  style={{
                    fontFamily: "'JetBrains Mono', monospace",
                    fontSize: 16,
                    fontWeight: 700,
                    color: pct >= 50 ? COLORS.green : COLORS.clay,
                  }}
                >
                  {pct}%
                </div>
              </div>
            );
          })}
        </div>
      )}

      <div style={{ textAlign: "center", marginTop: 24 }}>
        <button onClick={onGoExams} className="et-btn-interactive" style={btnGhost}>
          Go to exams
        </button>
      </div>
    </div>
  );
}

function DailyMenu({ submissions, onStart, onViewKey }) {
  const dates = listDailyDates();
  const today = todayStr();

  return (
    <div className="animate-fade-in">
      <div style={{ display: "grid", gap: 12 }}>
        {dates.map((date, idx) => {
          const key = `daily_${date}`;
          const sub = submissions[key];
          const isToday = date === today;
          const d = new Date(date + "T00:00:00");
          const niceDate = d.toLocaleDateString(undefined, {
            weekday: "short",
            month: "short",
            day: "numeric",
            year: "numeric",
          });
          const staggerClass = `stagger-${Math.min(idx + 1, 8)}`;

          return (
            <div
              key={date}
              className={`et-hover-card animate-slide-up ${staggerClass}`}
              style={{
                background: COLORS.inkSoft,
                border: `1px solid ${isToday ? COLORS.amber : COLORS.line}`,
                borderRadius: 8,
                padding: 18,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 12,
                flexWrap: "wrap",
              }}
            >
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <div
                    style={{
                      fontSize: 10,
                      letterSpacing: 2,
                      color: COLORS.amber,
                      fontFamily: "'JetBrains Mono', monospace",
                    }}
                  >
                    {date}
                  </div>
                  {isToday && (
                    <span
                      className="animate-pulse-glow"
                      style={{
                        fontSize: 10,
                        fontWeight: 700,
                        color: COLORS.ink,
                        background: COLORS.amber,
                        borderRadius: 10,
                        padding: "1px 8px",
                        letterSpacing: 0.5,
                      }}
                    >
                      TODAY
                    </span>
                  )}
                </div>
                <div
                  style={{
                    fontFamily: "'Spectral', serif",
                    fontSize: 17,
                    marginTop: 3,
                  }}
                >
                  {niceDate}
                </div>
                <div
                  style={{ fontSize: 12, color: COLORS.paperDim, marginTop: 3 }}
                >
                  {DAILY_COUNT} questions
                </div>
              </div>

              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                {sub && (
                  <div
                    style={{
                      fontFamily: "'JetBrains Mono', monospace",
                      fontSize: 15,
                      fontWeight: 700,
                      color: COLORS.green,
                    }}
                  >
                    {sub.score}/{sub.total}
                  </div>
                )}
                <button
                  onClick={() => (sub ? onViewKey(date) : onStart(date))}
                  className="et-btn-interactive"
                  style={sub ? btnGhost : btnPrimary}
                >
                  {sub ? "View key" : "Start"}
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function AnswersMenu({ submissions, now, onSelect, onGoExams }) {
  const anySubmitted = Object.values(submissions).some(Boolean);

  return (
    <div>
      <p
        style={{
          color: COLORS.paperDim,
          fontSize: 14,
          marginBottom: 24,
          lineHeight: 1.6,
        }}
      >
        Every round's answer key lives here. After you submit a round, the key
        becomes available with the correct answers displayed next to yours.
      </p>

      {!anySubmitted && (
        <div
          style={{
            textAlign: "center",
            padding: "40px 20px",
            border: `1px dashed ${COLORS.line}`,
            borderRadius: 8,
            color: COLORS.paperDim,
          }}
        >
          <p style={{ marginBottom: 16 }}>
            No answer keys yet — complete a round first.
          </p>
          <button onClick={onGoExams} style={btnGhost}>
            Go to exams
          </button>
        </div>
      )}

      <div style={{ display: "grid", gap: 12 }}>
        {Object.entries(QUESTIONS).map(([key, cat]) => {
          const sub = submissions[key];
          // Include assignments in the answers menu
          if (!sub) {
            return (
              <div
                key={key}
                style={{
                  background: COLORS.inkSoft,
                  border: `1px solid ${COLORS.line}`,
                  borderRadius: 8,
                  padding: 18,
                  opacity: 0.55,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <div>
                  <div
                    style={{
                      fontSize: 10,
                      letterSpacing: 2,
                      color: COLORS.paperDim,
                      fontFamily: "'JetBrains Mono', monospace",
                    }}
                  >
                    {cat.code}
                  </div>
                  <div
                    style={{
                      fontFamily: "'Spectral', serif",
                      fontSize: 18,
                      marginTop: 2,
                    }}
                  >
                    {cat.label}
                  </div>
                </div>
                <div style={{ fontSize: 12, color: COLORS.paperDim }}>
                  Not attempted
                </div>
              </div>
            );
          }

          const unlockAt = sub.submittedAt + REVEAL_MS;
          const remaining = unlockAt - now;
          const unlocked = remaining <= 0;

          return (
            <button
              key={key}
              onClick={() => onSelect(key)}
              style={{
                textAlign: "left",
                background: COLORS.inkSoft,
                border: `1px solid ${unlocked ? COLORS.green : COLORS.line}`,
                borderRadius: 8,
                padding: 18,
                cursor: "pointer",
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                width: "100%",
                fontFamily: "'Inter', sans-serif",
                color: COLORS.paper,
              }}
            >
              <div>
                <div
                  style={{
                    fontSize: 10,
                    letterSpacing: 2,
                    color: COLORS.amber,
                    fontFamily: "'JetBrains Mono', monospace",
                  }}
                >
                  {cat.code}
                </div>
                <div
                  style={{
                    fontFamily: "'Spectral', serif",
                    fontSize: 18,
                    marginTop: 2,
                  }}
                >
                  {cat.label}
                </div>
                <div
                  style={{
                    fontSize: 12,
                    color: COLORS.paperDim,
                    marginTop: 4,
                    fontFamily: "'JetBrains Mono', monospace",
                  }}
                >
                  Score {sub.score}/{sub.total}
                </div>
              </div>
              {unlocked ? (
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 6,
                    color: COLORS.green,
                    fontSize: 12,
                  }}
                >
                  <Unlock size={15} /> View key
                </div>
              ) : (
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "flex-end",
                    gap: 4,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      color: COLORS.amber,
                      fontSize: 12,
                    }}
                  >
                    <Lock size={13} /> Sealed
                  </div>
                  <div
                    style={{
                      fontSize: 11,
                      color: COLORS.paperDim,
                      fontFamily: "'JetBrains Mono', monospace",
                    }}
                  >
                    {formatCountdown(remaining)}
                  </div>
                </div>
              )}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function Home({ submissions, onStart, onViewKey }) {
  return (
    <div className="animate-fade-in">
      <div style={{ display: "grid", gap: 16 }}>
        {Object.entries(QUESTIONS)
          .filter(([, cat]) => !cat.isAssignment)
          .map(([key, cat], idx) => {
            const sub = submissions[key];
            const staggerClass = `stagger-${Math.min(idx + 1, 8)}`;
            return (
              <div
                key={key}
                className={`et-hover-card animate-slide-up ${staggerClass}`}
                style={{
                  background: COLORS.inkSoft,
                  border: `1px solid ${COLORS.line}`,
                  borderRadius: 8,
                  padding: 20,
                  position: "relative",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "flex-start",
                  }}
                >
                  <div>
                    <div
                      style={{
                        fontSize: 10,
                        letterSpacing: 2,
                        color: COLORS.amber,
                        fontFamily: "'JetBrains Mono', monospace",
                      }}
                    >
                      {cat.code}
                    </div>
                    <div
                      style={{
                        fontFamily: "'Spectral', serif",
                        fontSize: 24,
                        fontWeight: 600,
                        marginTop: 4,
                      }}
                    >
                      {cat.label}
                    </div>
                    <div
                      style={{
                        fontSize: 13,
                        color: COLORS.paperDim,
                        marginTop: 4,
                      }}
                    >
                      {cat.poolSize
                        ? `${cat.quizCount} questions`
                        : `${cat.items.length} questions`}
                    </div>
                  </div>
                  {sub && (
                    <div
                      style={{
                        textAlign: "right",
                        fontFamily: "'JetBrains Mono', monospace",
                      }}
                    >
                      <div
                        style={{
                          fontSize: 20,
                          fontWeight: 700,
                          color: COLORS.green,
                        }}
                      >
                        {sub.score}/{sub.total}
                      </div>
                      <div
                        style={{
                          fontSize: 10,
                          color: COLORS.paperDim,
                          letterSpacing: 1,
                        }}
                      >
                        SUBMITTED
                      </div>
                    </div>
                  )}
                </div>
                <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
                  <button
                    onClick={() => onStart(key)}
                    className="et-btn-interactive"
                    style={{
                      background: COLORS.paper,
                      color: COLORS.ink,
                      border: "none",
                      padding: "9px 16px",
                      borderRadius: 5,
                      fontSize: 13,
                      fontWeight: 600,
                      cursor: "pointer",
                      fontFamily: "'Inter', sans-serif",
                    }}
                  >
                    {sub ? "Retake" : "Start round"}
                  </button>
                  {sub && (
                    <button
                      onClick={() => onViewKey(key)}
                      className="et-btn-interactive"
                      style={{
                        background: "none",
                        color: COLORS.paperDim,
                        border: `1px solid ${COLORS.line}`,
                        padding: "9px 16px",
                        borderRadius: 5,
                        fontSize: 13,
                        cursor: "pointer",
                        fontFamily: "'Inter', sans-serif",
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                      }}
                    >
                      <Lock size={13} /> Answer key
                    </button>
                  )}
                </div>
              </div>
            );
          })}
      </div>
    </div>
  );
}

function CodeBlock({ code, language = "Python" }) {
  if (!code) return null;
  return (
    <div
      style={{
        background: COLORS.ink,
        border: `1px solid ${COLORS.line}`,
        borderRadius: 6,
        padding: "12px 14px",
        marginBottom: 16,
        overflowX: "auto",
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          marginBottom: 8,
          color: COLORS.paperDim,
        }}
      >
        <Code2 size={12} />
        <span
          style={{ fontSize: 10, letterSpacing: 1, textTransform: "uppercase" }}
        >
          {language}
        </span>
      </div>
      <pre
        style={{
          margin: 0,
          fontFamily: "'JetBrains Mono', monospace",
          fontSize: 13,
          color: COLORS.paper,
          whiteSpace: "pre-wrap",
          lineHeight: 1.5,
        }}
      >
        {code}
      </pre>
    </div>
  );
}

const TIER_LABEL = {
  easy: "Easy",
  medium: "Medium",
  hard: "Hard",
  zero: "Zero",
  beginner: "Beginner",
  intermediate: "Intermediate",
  advanced: "Advanced",
  pro: "Pro",
};

function Quiz({ items, meta, qIndex, answers, onSelect, onNext, onPrev }) {
  const [timeLeft, setTimeLeft] = useState(60);
  const onNextRef = useRef(onNext);
  const item = items[qIndex];
  const selected = answers[qIndex];
  const remainingCount = items.length - (qIndex + 1);
  const isAssignment = meta?.isAssignment;
  const hasAnswered = selected !== undefined;
  if (!item) return null;

  useEffect(() => {
    onNextRef.current = onNext;
  }, [onNext]);

  useEffect(() => {
    setTimeLeft(60);

    const timerId = window.setInterval(() => {
      setTimeLeft((prev) => {
        if (prev <= 1) {
          window.clearInterval(timerId);
          onNextRef.current();
          return 0;
        }
        return prev - 1;
      });
    }, 1000);

    return () => window.clearInterval(timerId);
  }, [qIndex, items.length]);

  const timeLabel = `${String(Math.floor(timeLeft / 60)).padStart(2, "0")}:${String(
    timeLeft % 60,
  ).padStart(2, "0")}`;

  return (
    <div className="animate-fade-in">
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 4,
          gap: 12,
        }}
      >
        <div
          style={{
            fontSize: 11,
            letterSpacing: 2,
            color: COLORS.amber,
            fontFamily: "'JetBrains Mono', monospace",
          }}
        >
          {meta.code}
          {item.tier && ` · ${TIER_LABEL[item.tier] || item.tier}`}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              padding: "6px 10px",
              borderRadius: 999,
              border: `1px solid ${COLORS.line}`,
              background: "transparent",
              color: COLORS.paperDim,
              fontFamily: "'JetBrains Mono', monospace",
              fontSize: 12,
            }}
          >
            <span>{remainingCount} remaining</span>
          </div>
          <div
            className={timeLeft <= 10 ? "animate-pulse-urgent" : ""}
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              padding: "6px 10px",
              borderRadius: 999,
              border: `1px solid ${timeLeft <= 10 ? COLORS.clay : COLORS.line}`,
              background:
                timeLeft <= 10
                  ? "rgba(193, 88, 63, 0.16)"
                  : "rgba(217, 164, 65, 0.12)",
              color: timeLeft <= 10 ? COLORS.clay : COLORS.amber,
              fontFamily: "'JetBrains Mono', monospace",
              fontSize: 12,
              transition: "all 0.3s ease",
            }}
          >
            <Clock size={13} />
            <span>{timeLabel}</span>
          </div>
        </div>
      </div>

      <div
        style={{
          height: 3,
          background: COLORS.line,
          borderRadius: 2,
          marginTop: 10,
          marginBottom: 24,
          overflow: "hidden",
        }}
      >
        <div
          style={{
            height: "100%",
            width: `${((qIndex + 1) / items.length) * 100}%`,
            background: COLORS.amber,
            transition: "width 0.4s cubic-bezier(0.4, 0, 0.2, 1)",
          }}
        />
      </div>

      <div
        key={qIndex}
        className="animate-slide-up"
        style={{
          background: COLORS.inkSoft,
          border: `1px solid ${COLORS.line}`,
          borderRadius: 8,
          padding: 24,
        }}
      >
        <div
          style={{
            fontFamily: "'Spectral', serif",
            fontSize: 19,
            fontWeight: 600,
            lineHeight: 1.5,
            marginBottom: item.code ? 14 : 20,
          }}
        >
          {`Q${qIndex + 1}. ${item.q}`}
        </div>
        <CodeBlock code={item.code} language={item.srcLabel?.includes("Java") ? "Java" : item.srcLabel?.includes("C") ? "C" : meta?.label?.includes("Java") ? "Java" : meta?.label?.includes("C") ? "C" : "Python"} />
        <div style={{ display: "grid", gap: 10 }}>
          {item.opts.map((opt, i) => {
            const isSelected = selected === i;
            const isDisabled = isAssignment && hasAnswered;
            return (
              <button
                key={i}
                onClick={() => !isDisabled && onSelect(i)}
                disabled={isDisabled}
                className={`et-option-btn ${isSelected ? "selected" : ""}`}
                style={{
                  textAlign: "left",
                  padding: "12px 14px",
                  borderRadius: 6,
                  border: `1px solid ${isSelected ? COLORS.amber : COLORS.line}`,
                  background: isSelected
                    ? "rgba(217,164,65,0.12)"
                    : "transparent",
                  color: COLORS.paper,
                  fontSize: 14,
                  cursor: isDisabled ? "not-allowed" : "pointer",
                  fontFamily: "'Inter', sans-serif",
                  display: "flex",
                  alignItems: "center",
                  gap: 10,
                  opacity: isDisabled && !isSelected ? 0.5 : 1,
                }}
              >
                <span
                  style={{
                    width: 20,
                    height: 20,
                    borderRadius: "50%",
                    border: `1px solid ${isSelected ? COLORS.amber : COLORS.paperDim}`,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 11,
                    flexShrink: 0,
                    color: isSelected ? COLORS.amber : COLORS.paperDim,
                    transition: "all 0.2s ease",
                  }}
                >
                  {String.fromCharCode(65 + i)}
                </span>
                {opt}
              </button>
            );
          })}
        </div>
      </div>

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          marginTop: 20,
        }}
      >
        <button
          onClick={onPrev}
          disabled={qIndex === 0}
          className="et-btn-interactive"
          style={{
            background: "none",
            border: `1px solid ${COLORS.line}`,
            color: qIndex === 0 ? COLORS.line : COLORS.paperDim,
            padding: "9px 16px",
            borderRadius: 5,
            fontSize: 13,
            cursor: qIndex === 0 ? "default" : "pointer",
            display: "flex",
            alignItems: "center",
            gap: 6,
          }}
        >
          <ArrowLeft size={14} /> Back
        </button>
        <button
          onClick={onNext}
          disabled={selected === undefined}
          className="et-btn-interactive"
          style={{
            background: selected === undefined ? COLORS.line : COLORS.paper,
            color: selected === undefined ? COLORS.paperDim : COLORS.ink,
            border: "none",
            padding: "9px 18px",
            borderRadius: 5,
            fontSize: 13,
            fontWeight: 600,
            cursor: selected === undefined ? "default" : "pointer",
            display: "flex",
            alignItems: "center",
            gap: 6,
          }}
        >
          {qIndex === items.length - 1 ? "Submit" : "Next"}{" "}
          <ArrowRight size={14} />
        </button>
      </div>
    </div>
  );
}

function Result({
  category,
  record,
  leaderboard,
  currentUser,
  onViewKey,
  onHome,
  onRetake,
}) {
  const cat = getCategoryMeta(category, currentUser);
  if (!record) return null;
  const pct = Math.round((record.score / record.total) * 100);
  const items = record.items || cat.items;

  const myRank =
    leaderboard && leaderboard.length
      ? leaderboard.findIndex((r) => r.username === currentUser) + 1
      : 0;
  const totalMembers = leaderboard ? leaderboard.length : 0;

  return (
    <div className="animate-fade-in" style={{ padding: "20px 0" }}>
      <div style={{ textAlign: "center" }}>
        <div
          style={{
            fontSize: 11,
            letterSpacing: 2,
            color: COLORS.amber,
            fontFamily: "'JetBrains Mono', monospace",
          }}
        >
          {cat.code} — RESULT
        </div>
        <div
          className="animate-score"
          style={{
            margin: "24px auto",
            width: 160,
            height: 160,
            borderRadius: "50%",
            border: `3px solid ${pct >= 50 ? COLORS.green : COLORS.clay}`,
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            boxShadow: `0 0 24px ${pct >= 50 ? "rgba(63,143,99,0.3)" : "rgba(193,88,63,0.3)"}`,
          }}
        >
          <div
            style={{
              fontFamily: "'JetBrains Mono', monospace",
              fontSize: 36,
              fontWeight: 700,
              color: pct >= 50 ? COLORS.green : COLORS.clay,
            }}
          >
            {record.score}/{record.total}
          </div>
          <div style={{ fontSize: 12, color: COLORS.paperDim, marginTop: 4 }}>
            {pct}%
          </div>
        </div>
        <div
          style={{
            fontFamily: "'Spectral', serif",
            fontSize: 20,
            marginBottom: 6,
          }}
        >
          {cat.label} round complete
        </div>

        {myRank > 0 && (
          <div
            className="animate-pulse-glow"
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 8,
              background: COLORS.inkSoft,
              border: `1px solid ${COLORS.line}`,
              borderRadius: 20,
              padding: "6px 16px",
              margin: "6px 0 14px",
            }}
          >
            <Trophy size={14} color={RANK_MEDAL[myRank] || COLORS.amber} />
            <span
              style={{
                fontSize: 13,
                fontFamily: "'JetBrains Mono', monospace",
              }}
            >
              Rank #{myRank}{" "}
              <span style={{ color: COLORS.paperDim }}>
                of {totalMembers} overall
              </span>
            </span>
          </div>
        )}

        <p
          style={{
            color: COLORS.paperDim,
            fontSize: 13,
            maxWidth: 420,
            margin: "0 auto 24px",
            lineHeight: 1.6,
          }}
        >
          Your score is final. The detailed answer key shows the correct answers
          alongside your own responses.
        </p>
        <div
          style={{
            display: "flex",
            gap: 10,
            justifyContent: "center",
            flexWrap: "wrap",
            marginBottom: 32,
          }}
        >
          <button onClick={onViewKey} className="et-btn-interactive" style={btnPrimary}>
            <Unlock size={13} style={{ marginRight: 6 }} /> View answer key
          </button>
          <button onClick={onRetake} className="et-btn-interactive" style={btnGhost}>
            <RotateCcw size={13} style={{ marginRight: 6 }} /> Retake
          </button>
          <button onClick={onHome} className="et-btn-interactive" style={btnGhost}>
            All exams
          </button>
        </div>
      </div>

      <div
        style={{
          fontSize: 11,
          letterSpacing: 2,
          color: COLORS.paperDim,
          textTransform: "uppercase",
          textAlign: "center",
          marginBottom: 16,
        }}
      >
        Question review
      </div>
      <div style={{ position: "relative" }}>
        <div
          style={{
            position: "absolute",
            left: "50%",
            top: 0,
            bottom: 0,
            width: 1,
            background: COLORS.line,
            transform: "translateX(-50%)",
          }}
        />
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {items.map((item, i) => {
            const userAns = record.answers[i];
            const isCorrect = userAns === item.correct;
            const fromLeft = i % 2 === 0;
            const staggerClass = `stagger-${Math.min((i % 8) + 1, 8)}`;
            return (
              <div
                key={i}
                style={{
                  display: "flex",
                  justifyContent: fromLeft ? "flex-start" : "flex-end",
                  width: "100%",
                }}
              >
                <div
                  className={`et-hover-card animate-slide-up ${staggerClass}`}
                  style={{
                    width: "82%",
                    maxWidth: 460,
                    background: COLORS.inkSoft,
                    border: `1px solid ${isCorrect ? "rgba(63,143,99,0.4)" : "rgba(193,88,63,0.4)"}`,
                    borderRadius: 8,
                    padding: 14,
                    position: "relative",
                  }}
                >
                  <div
                    style={{
                      position: "absolute",
                      top: 16,
                      [fromLeft ? "right" : "left"]: -19,
                      width: 10,
                      height: 10,
                      borderRadius: "50%",
                      background: isCorrect ? COLORS.green : COLORS.clay,
                      border: `2px solid ${COLORS.ink}`,
                      display: window.innerWidth < 640 ? "none" : "block",
                    }}
                  />
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      gap: 10,
                    }}
                  >
                    <div
                      style={{ fontSize: 13, fontWeight: 500, lineHeight: 1.5 }}
                    >
                      {`Q${i + 1}. ${item.q}`}
                    </div>
                    {isCorrect ? (
                      <CheckCircle2
                        size={16}
                        color={COLORS.green}
                        style={{ flexShrink: 0 }}
                      />
                    ) : (
                      <XCircle
                        size={16}
                        color={COLORS.clay}
                        style={{ flexShrink: 0 }}
                      />
                    )}
                  </div>
                  {item.code && (
                    <div style={{ marginTop: 8 }}>
                      <CodeBlock code={item.code} language={item.srcLabel?.includes("Java") ? "Java" : item.srcLabel?.includes("C") ? "C" : meta?.label?.includes("Java") ? "Java" : meta?.label?.includes("C") ? "C" : "Python"} />
                    </div>
                  )}
                  <div style={{ marginTop: 8, fontSize: 12 }}>
                    <div style={{ color: COLORS.green }}>
                      Correct: {item.opts[item.correct]}
                    </div>
                    {!isCorrect && (
                      <div style={{ color: COLORS.clay, marginTop: 2 }}>
                        Your answer:{" "}
                        {userAns !== undefined
                          ? item.opts[userAns]
                          : "Not answered"}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

const btnPrimary = {
  background: COLORS.paper,
  color: COLORS.ink,
  border: "none",
  padding: "10px 18px",
  borderRadius: 5,
  fontSize: 13,
  fontWeight: 600,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  transition: "all 0.2s cubic-bezier(0.16, 1, 0.3, 1)",
};

const btnGhost = {
  background: "none",
  color: COLORS.paperDim,
  border: `1px solid ${COLORS.line}`,
  padding: "10px 18px",
  borderRadius: 5,
  fontSize: 13,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  transition: "all 0.2s cubic-bezier(0.16, 1, 0.3, 1)",
};

function AnswerKey({ category, record, now, onHome, onRetake, currentUser }) {
  const cat = getCategoryMeta(category, currentUser);
  const keyItems = (record && record.items) || cat.items;

  if (!record) {
    return (
      <div
        style={{
          textAlign: "center",
          padding: "40px 0",
          color: COLORS.paperDim,
        }}
      >
        <p>No submission yet for {cat.label}. Complete the round first.</p>
        <button onClick={onHome} style={{ ...btnGhost, marginTop: 16 }}>
          Back to exams
        </button>
      </div>
    );
  }

  // Only show questions the user actually submitted (answered)
  const answeredIndices = Object.keys(record.answers).map(Number);
  const submittedCount = answeredIndices.length;
  const correctCount = answeredIndices.filter(
    (i) => record.answers[i] === keyItems[i]?.correct,
  ).length;
  const pct =
    submittedCount > 0 ? Math.round((correctCount / submittedCount) * 100) : 0;

  return (
    <div>
      {/* Score summary header */}
      <div
        style={{
          background: COLORS.inkSoft,
          border: `1px solid ${COLORS.line}`,
          borderRadius: 10,
          padding: "18px 22px",
          marginBottom: 22,
          display: "flex",
          alignItems: "center",
          gap: 20,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            width: 64,
            height: 64,
            borderRadius: "50%",
            border: `3px solid ${pct >= 50 ? COLORS.green : COLORS.clay}`,
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          <div
            style={{
              fontFamily: "'JetBrains Mono', monospace",
              fontSize: 18,
              fontWeight: 700,
              color: pct >= 50 ? COLORS.green : COLORS.clay,
              lineHeight: 1,
            }}
          >
            {pct}%
          </div>
        </div>
        <div style={{ flex: 1, minWidth: 160 }}>
          <div
            style={{
              fontFamily: "'Spectral', serif",
              fontSize: 18,
              fontWeight: 600,
              marginBottom: 4,
            }}
          >
            {cat.label} — Answer Key
          </div>
          <div
            style={{
              display: "flex",
              gap: 18,
              fontSize: 13,
              flexWrap: "wrap",
            }}
          >
            <span>
              <span style={{ color: COLORS.paperDim }}>Submitted: </span>
              <span
                style={{
                  fontFamily: "'JetBrains Mono', monospace",
                  color: COLORS.paper,
                }}
              >
                {submittedCount}/{record.total}
              </span>
            </span>
            <span>
              <span style={{ color: COLORS.paperDim }}>Correct: </span>
              <span
                style={{
                  fontFamily: "'JetBrains Mono', monospace",
                  color: COLORS.green,
                }}
              >
                {correctCount}
              </span>
            </span>
            <span>
              <span style={{ color: COLORS.paperDim }}>Wrong: </span>
              <span
                style={{
                  fontFamily: "'JetBrains Mono', monospace",
                  color: COLORS.clay,
                }}
              >
                {submittedCount - correctCount}
              </span>
            </span>
            <span>
              <span style={{ color: COLORS.paperDim }}>Skipped: </span>
              <span
                style={{
                  fontFamily: "'JetBrains Mono', monospace",
                  color: COLORS.paperDim,
                }}
              >
                {record.total - submittedCount}
              </span>
            </span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {onRetake && (
            <button
              onClick={onRetake}
              style={{
                ...btnGhost,
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
            >
              <RotateCcw size={13} /> Retake
            </button>
          )}
          <button onClick={onHome} style={btnGhost}>
            All exams
          </button>
        </div>
      </div>

      {/* Only the answered questions */}
      <div
        style={{
          fontSize: 11,
          letterSpacing: 2,
          color: COLORS.paperDim,
          textTransform: "uppercase",
          marginBottom: 14,
        }}
      >
        {submittedCount} question{submittedCount !== 1 ? "s" : ""} submitted
      </div>

      <div style={{ display: "grid", gap: 14 }}>
        {answeredIndices.map((i) => {
          const item = keyItems[i];
          if (!item) return null;
          const userAns = record.answers[i];
          const isCorrect = userAns === item.correct;
          return (
            <div
              key={i}
              style={{
                background: COLORS.inkSoft,
                border: `1px solid ${
                  isCorrect ? "rgba(63,143,99,0.45)" : "rgba(193,88,63,0.45)"
                }`,
                borderRadius: 8,
                padding: 16,
              }}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  gap: 10,
                }}
              >
                <div style={{ fontSize: 14, fontWeight: 500, lineHeight: 1.5 }}>
                  {`Q${i + 1}. ${item.q}`}
                </div>
                {isCorrect ? (
                  <CheckCircle2
                    size={18}
                    color={COLORS.green}
                    style={{ flexShrink: 0 }}
                  />
                ) : (
                  <XCircle
                    size={18}
                    color={COLORS.clay}
                    style={{ flexShrink: 0 }}
                  />
                )}
              </div>
              {item.code && (
                <div style={{ marginTop: 10 }}>
                  <CodeBlock code={item.code} language={item.srcLabel?.includes("Java") ? "Java" : item.srcLabel?.includes("C") ? "C" : meta?.label?.includes("Java") ? "Java" : meta?.label?.includes("C") ? "C" : "Python"} />
                </div>
              )}
              <div style={{ marginTop: item.code ? 0 : 10, fontSize: 13 }}>
                <div style={{ color: COLORS.green }}>
                  Correct: {item.opts[item.correct]}
                </div>
                {!isCorrect && (
                  <div style={{ color: COLORS.clay, marginTop: 2 }}>
                    Your answer: {item.opts[userAns]}
                  </div>
                )}
                {item.explanation && (
                  <div style={{ marginTop: 12, paddingTop: 12, borderTop: `1px solid ${COLORS.line}` }}>
                    <div style={{ fontSize: 11, letterSpacing: 1, textTransform: "uppercase", color: COLORS.paperDim, marginBottom: 6 }}>
                      Explanation
                    </div>
                    <div style={{ fontSize: 13, lineHeight: 1.6, color: COLORS.paper }}>
                      {item.explanation}
                    </div>
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>

      <div style={{ textAlign: "center", marginTop: 24 }}>
        {onRetake && (
          <button
            onClick={onRetake}
            style={{
              ...btnGhost,
              marginRight: 10,
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            <RotateCcw size={13} /> Retake
          </button>
        )}
        <button onClick={onHome} style={btnGhost}>
          Back to exams
        </button>
      </div>
    </div>
  );
}
