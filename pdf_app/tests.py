from django.test import TestCase

# Create your tests here.

def decorator_function(func):
    print('something to happen before calling function')

    print('AFTER THE FUNCTION')
    return func

@decorator_function
def real_function():
    print("real function")


real_function()
