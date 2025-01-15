def example_function(*args):
    print(args)
    print(type(args))

example_function(1, 2, 3, 4, 5)  # (1, 2, 3, 4, 5)
example_function("ddwq",3321,(22,3),[1,2,3,4])


def greet(message, *args):
    print(message)
    for name in args:
        print(f"Hello, {name}!")

greet("Welcome", "Alice", "Bob", "Charlie")




def sum_numbers(*args):
    return sum(args)

numbers = (1, 2, 3, 4)
print(sum_numbers(*numbers))  # 10


class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

    def greet(self):
        return f"Hello, my name is {self.name}."

    def age_next_year(self):
        nextage = self.age + 1
        return f"I will be {nextage} next year"


    def full_intro(self):
        greeting = self.greet()  # 调用另一个方法
        aggg = self.age_next_year()
        return f"{greeting}{aggg}"

person = Person("Alice", 30)
print(person.full_intro())
