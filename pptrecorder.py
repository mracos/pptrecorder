import argparse, io, logging
from time import time
from types import MethodType
from multiprocessing import Queue, Process

import pyscreenshot
from pptx import Presentation
from pptx.util import Inches

log = logging.getLogger()

def parse_args():
    parser = argparse.ArgumentParser(
        description=("Record a screencast as a"
                     " series of power-point slides")
    )

    parser.add_argument("file", metavar="file", type=str)
    parser.add_argument("-v", "--verbosity", action="store_true")
    return parser.parse_args()

def append_screenshot_queue(image_queue):
    """
        append a screenshot of the actual screen
        to a FIFO queue
    """
    logging.info("taking screenshot")
    image = pyscreenshot.grab()
    image_queue.put(image)

def add_slide(ppt, queue_image):
    """
        append an slide to the presentation with
        the image get from the FIFO queue
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

    logging.info("adding image to presentation")
    pic = slide.shapes.add_picture(
        bimage, left, top, height=height
    )

def record_screen():
    """
        creates the pptx in memory and add the slides
        with the screenshot
    """
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

    logging.info(("starting get_image and add_slide "
               "process"))
    process_get_image.start()
    process_add_slide.start()

    try:
        while True:
            process_get_image.run()
            process_add_slide.run()
    except KeyboardInterrupt:
        logging.info(("terminating add_slide ,"
                   "get_image and queue process"))
        process_add_slide.terminate()
        process_get_image.terminate()
        image_queue.terminate()
    finally:
        return ppt

def main():
    args = parse_args()
    log.setLevel(logging.INFO if args.verbosity else logging.ERROR)

    print("starting to record")
    time_start = time()
    ppt = record_screen()
    time_elapsed = (time() - time_start)
    print("recorded {0:.3f} seconds".format(time_elapsed))
    print("saving to file")

    ppt.save(args.file)

if __name__ == "__main__":
    main()
