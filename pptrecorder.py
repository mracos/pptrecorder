import argparse, io
from types import MethodType
from multiprocessing import Queue, Process

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

def append_screenshot_queue(image_queue):
    print("taking screenshot")
    image = pyscreenshot.grab()
    image_queue.put(image)

def add_slide(ppt, queue_image):
    """
        append an slide to the presentation with
        the current image
    """
    image = queue_image.get()
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
    image_queue = Queue()
    process_get_image = Process(
        target=append_screenshot_queue,
        args=(image_queue,)
    )
    process_add_slide = Process(
        target=add_slide,
        args=(ppt, image_queue)
    )

    process_get_image.start()
    process_add_slide.start()

    try:
        while True:
            process_get_image.run()
            process_add_slide.run()
    except KeyboardInterrupt:
        process_add_slide.terminate()
        process_get_image.terminate()
        image_queue.terminate()
    finally:
        return ppt

def main():
    args = parse_args()
    ppt = record_screen()
    print("saving to file")
    ppt.save(args.file)

if __name__ == '__main__':
    main()
