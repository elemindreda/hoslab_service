import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
from PIL import ImageOps
(width, height) = (160, 35)


getimg = Image.open('HOSLAB-LOGO-HQ.jpg')
img = getimg.resize((160, 35), Image.ANTIALIAS)
logo_hoslab = np.array(img)

test = logo_hoslab.ravel()

imgplot = plt.imshow(test.reshape(height, width, 3))