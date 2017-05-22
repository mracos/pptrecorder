import argparse, io
from types import MethodType

import pyscreenshot
from pptx import Presentation
from pptx.util import Inches


def parse_args():
    parser = argparse.ArgumentParser(
        description=("Record a screencast as a"
                     " series of power-point slides")
    )

    parser.add_argument("file", metavar="file", type=str)
    return parser.parse_args()

def get_actual_image():
    image = pyscreenshot.grab()
    return image

def add_slide(ppt, image):
    """ append an slide to the presentation with
        the current image """
    blank_slide_layout = ppt.slide_layouts[6]

    slide = ppt.slides.add_slide(blank_slide_layout)
    left = top = Inches(0)
    height = Inches(10)

    # workaround because the add_picture expects an read()
    # to return the bytes
    bimage = io.BytesIO()
    image.save(bimage, format='PNG')
    bimage.read = MethodType(lambda self: self.getvalue(), bimage)

    pic = slide.shapes.add_picture(
        bimage, left, top, height=height
    )

def record_screen():
    ppt = Presentation()
    try:
        while True:
            print("adding image")
            image = get_actual_image()
            add_slide(ppt, image)
    except KeyboardInterrupt:
        print("saving to file")
        pass
    finally:
        return ppt

def main():
    args = parse_args()
    ppt = record_screen()
    ppt.save(args.file)

if __name__ == '__main__':
    main()
