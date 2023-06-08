import math
import statistics


def avg_generator(data, numexpts):
    trimmed_data = [[],[],[],[],[],[]], [[],[],[],[],[],[]]
    # Find the minimum and maximum values from the first elements of each set

    #if the initial potential is lower than the final potential / becomes more negative
    if data[0][1][0] > data[0][1][-1]:
        in_val = max(data[i][1][0] for i in range(numexpts))
        fin_val = max(data[i][1][-1] for i in range(numexpts))
        for i in range(numexpts):
            set_len = len(data[i][1])
            for x in range(set_len):
                if data[i][1][x] < (in_val*0.99) and data[i][1][x] > (fin_val* 1.01):
                    trimmed_data[1][i].append(data[i][1][x])
                    trimmed_data[0][i].append(data[i][2][x])

    #if the initial potential is higher than the final potential / becomes more positive
    else:
        in_val = min(data[i][1][0] for i in range(numexpts))
        fin_val = min(data[i][1][-1] for i in range(numexpts))
        for i in range(numexpts):
            set_len = len(data[i][1])
            for x in range(set_len):
                if data[i][1][x] > (in_val* 1.01) and data[i][1][x] < (fin_val* 0.99):
                    trimmed_data[1][i].append(data[i][1][x])
                    trimmed_data[0][i].append(data[i][2][x])


    # Create new trimmed sets

    # Get the number of values in each trimmed set
    num_values = len(trimmed_data[0][0])

    # Calculate the standard error and standard deviation for each point in the trimmed sets
    #[0: avg_pots], [1: avg_curs], [2: std_devs], [3: std_errors], [4: avg_min_stdev], [5: avg_plus_stdev], [6:avg_min_stderror],
    # [7:avg_plus_stderror], [8: set_std_errors], [9: set_std_error]
    stats_data = [],[],[],[],[],[],[],[],[]
    for i in range(num_values):
        curs_list = []
        sum_pots = 0
        sum_curs = 0
        for x in range(numexpts):
            sum_pots += trimmed_data[1][x][i]
            sum_curs += trimmed_data[0][x][i]
            curs_list.append(trimmed_data[0][x][i])
        avg_dev_pots = sum_pots / numexpts
        avg_dev_curs = sum_curs / numexpts
        stats_data[0].append(avg_dev_pots)
        stats_data[1].append(avg_dev_curs)
        std_dev = statistics.stdev(curs_list)
        stats_data[2].append(std_dev)
        std_error = std_dev / math.sqrt(numexpts)
        stats_data[3].append(std_error)
        stats_data[4].append(avg_dev_curs - std_dev)
        stats_data[5].append(avg_dev_curs + std_dev)
        stats_data[6].append(avg_dev_curs - std_error)
        stats_data[7].append(avg_dev_curs + std_error)
    print('stats', stats_data)
    return stats_data


def rsq(logs, pots):
    sqx_list = []
    sqy_list = []
    xy_list = []
    nums = 0
    for log in logs:
        sqy_list.append(log**2)
        nums += 1
    for pot in pots:
        sqx_list.append(pot**2)
    for log in logs:
        ind = logs.index(log)
        xy_list.append(log*(pots[ind]))
    sum_xy = sum(xy_list)
    r_squared = abs((nums * sum_xy - (sum(pots)) * (sum(logs))) / (((math.sqrt(nums * sum(sqx_list) - (sum(pots)**2))) * (math.sqrt(nums * sum(sqy_list) - (sum(logs)**2))))))**2
    return r_squared


def get_data(expt):
#get potential, current, current density, and log values for each experiment set
    uploadedData = open(expt)
    dataLines = uploadedData.readlines()
    pot_list = []
    cur_list = []
    curdens_list = []
    log_list = []
    del dataLines[0]
    num=0
    for line in dataLines:
        num += 1
        line_temp = line.strip().split(",")
        pot_list.append(float(line_temp[0]))
        cur_list.append(float(line_temp[1]))
        curdens_list.append((float(line_temp[1]) / 0.196)* 1000)
        log_list.append(math.log10(abs(float(line_temp[1]))))

    if cur_list[-4] < 0 > cur_list[4]:
        OERHER = 'HER'
    elif cur_list[-4] > 0 < cur_list[4]:
        OERHER = 'OER'
    else:
        OERHER = 'unknown'
    print(curdens_list, '\n', cur_list )
    print('logs: ', log_list)
    print('potentials: ', pot_list)
    return cur_list, pot_list, curdens_list, log_list, OERHER

def getTafel(logs, pots):
#step1: determine 5% of domain (minimum domain). This is the minimum set size.
    min_domain = int(round((len(pots) * .05), 0))
#This is how many values are in the set. -1 to account for ts use as in indexes.
    pots_len = len(pots)-1
    print(f'pots len{pots_len}')
    # the program will stop testing sets once num <= test_stop.
    test_stop = pots_len - min_domain
#step2: get the slope of each possible set containing at least 20% of domain values.
    slope_list = [],[],[],[]
    #num1 is the index of the initial set value and num2 is the final.
    num1 = 0
    #This is the first x1 value.
    while num1 <= test_stop:
        num2 = min_domain + num1
        while num2 <= pots_len:
            r_value = rsq(logs[num1:num2+1],pots[num1:num2+1])
            slope = abs((logs[num2]) - (logs[num1])) / ((pots[num2]) - (pots[num1]))
            if abs(r_value) > 0.9725:
                slope_list[0].append(abs(slope))
                slope_list[1].append(abs(r_value))
                slope_list[2].append(num1)
                slope_list[3].append(num2)
            num2 += 1
        num1 += 1
#step3: find slopes close to the max slope.
    max_slope = max(slope_list[0])
    max_ind = slope_list[0].index(max_slope)
    max_r2 = slope_list[1][max_ind]
    slope_list1 = [],[],[],[],[]
    num = 0
    for slo in slope_list[0]:
        slo = abs(slo)
        slor2 = slope_list[1][num]
        num += 1
        if (slo > (max_slope * 0.85)) and (slor2 > (max_r2 * 0.9985)):
            sloind = slope_list[0].index(slo)
            r_2_val = slope_list[1][sloind]
            low = slope_list[2][sloind]
            high = slope_list[3][sloind]
            slope_list1[1].append(r_2_val)
            slope_list1[2].append(low)
            slope_list1[3].append(high)
            slope_list1[4].append((logs[high])-(logs[low]))
#Step 4: select set with greatest range.
    max_range = max(slope_list1[4])
#get variables to be returned.
    max_index = slope_list1[4].index(max_range)
    log_ind1 = slope_list1[2][max_index]
    log_ind2 = slope_list1[3][max_index]
    r_2 = rsq(pots[log_ind1:log_ind2+1],logs[log_ind1:log_ind2+1])
    print(f'r: {r_2}')
    log_low = logs[0]
    log_high = logs[-1]
    m_slope = (logs[log_ind2]- logs[log_ind1])/((pots[log_ind2])-(pots[log_ind1]))
    Tslope = ((1 / m_slope ) * 1000) / 2.303
    print(f"slope: {Tslope}")
    Tslope = f'{((1 / m_slope ) * 1000) / 2.303:.2f}'
    return m_slope, Tslope, max_range, r_2, log_ind1, log_ind2, log_low, log_high

