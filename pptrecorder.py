import argparse, io, logging, signal
from time import time
from types import MethodType
from multiprocessing import Queue, Process

import pyscreenshot
from pptx import Presentation
from pptx.util import Emu
from PIL import Image

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
    while True:
        logging.info(
            "taking screenshot, image_queue: [{}]".format(image_queue.qsize())
        )
        image = pyscreenshot.grab()
        image_queue.put(image)

def resize_patch_image(image, size):
    """
        gets the image from the queue
        and it returns it resized and 'patched'
        to be processed bu the add_picture()
    """
    image.thumbnail(size)
    # workaround because the add_picture expects an read()
    # to return the bytes
    b_image = io.BytesIO()
    image.save(b_image, format='PNG')
    b_image.read = MethodType(lambda self: self.getvalue(), b_image)
    return b_image

def add_slide_to_ppt(ppt, image_queue):
    """
        append an slide to the presentation with
        the image get from the FIFO queue
    """
    size = (
        Emu(ppt.slide_height).pt,
        Emu(ppt.slide_width).pt
    )
    blank_slide_layout = ppt.slide_layouts[6]
    left = top = Emu(0)
    while True:
        image = resize_patch_image(image_queue.get(), size)
        slide = ppt.slides.add_slide(blank_slide_layout)

        logging.info(
            "adding image to presentation, slides: [{}]".format(len(ppt.slides)),
        ) 
        pic = slide.shapes.add_picture(
            image, left, top
        )

def record_screen():
    """
        creates the pptx in memory and add the slides
        with the screenshot
    """
    ppt = Presentation()

    # make child processes ignore ctrl-c
    signal.pthread_sigmask(signal.SIG_BLOCK,
                           (signal.SIGINT,))
    image_queue = Queue()
    process_get_image = Process(
        target=append_screenshot_queue,
        args=(image_queue,)
    )
    process_add_slide = Process(
        target=add_slide_to_ppt,
        args=(ppt, image_queue)
    )

    process_add_slide.start()
    process_get_image.start()

    # now we catch CTRL-C again
    signal.pthread_sigmask(signal.SIG_UNBLOCK,
                           (signal.SIGINT,))

    logging.info(
        ("started processes: get_image "
         "and add_slide with PIDs: {}, {}").format(
             process_get_image.pid, process_add_slide.pid
         )
    )

    try:
        process_add_slide.join()
        process_get_image.join()
    except KeyboardInterrupt:
        logging.info(
            "terminate get_image and add_slide with PIDs: {}, {}"
                .format(process_get_image.pid, process_add_slide.pid)
        )
        process_add_slide.terminate()
        process_get_image.terminate()
        image_queue.close()
    finally:
        return ppt

if __name__ == "__main__":
    args = parse_args()
    log.setLevel(logging.INFO if args.verbosity else logging.ERROR)

    print("starting to record")
    time_start = time()
    ppt = record_screen()
    time_elapsed = (time() - time_start)
    print("recorded {0:.3f} seconds".format(time_elapsed))
    print("saving to file")

    ppt.save(args.file)
