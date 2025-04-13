class Point:
    __match_args__ = ('x', 'y')
    def __init__(self, x, y):
        self.x = x
        self.y = y


var = 42

#point = Point(1, var)
#Point(1, y=var)
#Point(x=1, y=var)
point =  Point(0, var)



match point:
    case Point(x=0, y=0):
        print("Origin")
    case Point(0, y=y):
        print(f"Y={y}")
    case Point(x=x, y=0):
        print(f"X={x}")
    case Point():
        print("Somewhere else")
    case _:
        print("Not a point")


