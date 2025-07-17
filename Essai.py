import math

def area_of_circle(radius):
    """Calculate the area of a circle given its radius."""
    return math.pi * radius ** 2

# Example usage
if __name__ == "__main__":
    r = float(input("Enter the radius of the circle: "))
    area = area_of_circle(r)
    print(f"The area of the circle with radius {r} is {area:.2f}")
