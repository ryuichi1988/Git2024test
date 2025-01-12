import cv2

# 加载图像
image_path = r'C:\1\1.jpg'
image = cv2.imread(image_path)

# 定义鼠标回调函数
def draw_rectangle(event, x, y, flags, param):
    global x_start, y_start, drawing, img_copy
    if event == cv2.EVENT_LBUTTONDOWN:  # 鼠标左键按下
        drawing = True
        x_start, y_start = x, y
    elif event == cv2.EVENT_MOUSEMOVE:  # 鼠标移动
        if drawing:
            img_copy = image.copy()
            cv2.rectangle(img_copy, (x_start, y_start), (x, y), (0, 255, 0), 2)
            cv2.imshow("Select Region", img_copy)
    elif event == cv2.EVENT_LBUTTONUP:  # 鼠标左键释放
        drawing = False
        print(f"Region: ({x_start}, {y_start}, {x}, {y})")  # 输出矩形区域坐标

# 初始化
drawing = False
x_start, y_start = -1, -1
img_copy = image.copy()

# 创建窗口并绑定回调函数
cv2.namedWindow("Select Region")
cv2.setMouseCallback("Select Region", draw_rectangle)

# 显示图像
while True:
    cv2.imshow("Select Region", img_copy)
    key = cv2.waitKey(1) & 0xFF
    if key == 27:  # 按 ESC 键退出
        break

cv2.destroyAllWindows()
