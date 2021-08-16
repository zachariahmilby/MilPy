import multiprocessing as mp


def get_appropriate_number_of_cores():
    n_cores = int(mp.cpu_count()/2 - 1)
    if n_cores < 1:
        n_cores = 1
    return n_cores


def set_processor_pool(n_cores):
    """
    mp.get_context('fork') allows parallel processing without 'if __name__ == '__main__'.
    """
    return mp.get_context('fork').Pool(n_cores)


def get_multiprocessing_pool():
    n_cores = get_appropriate_number_of_cores()
    pool = set_processor_pool(n_cores)
    return pool


def cleanup_parallel_processing(pool):
    """
    No one knows what these do...but things don't work without them.
    """
    pool.close()
    pool.join()
