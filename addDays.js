Date.prototype.addDays = function (s) {

    var targetDays = parseInt(s)
    var thisYear = parseInt(this.getFullYear())
    var thisDays = parseInt(this.getDate())
    var thisMonth = parseInt(this.getMonth() + 1)

    var currDays = thisDays;
    var currMonth = thisMonth;
    var currYear = thisYear;

    var monthArr;

    var nonleap = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    // leap year  
    var leap = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

    if ((thisYear % 4) == 0) {
        if ((thisYear % 100) == 0 && (thisYear % 400) != 0) { monthArr = nonleap; }
        else { monthArr = leap; }
    }
    else { monthArr = nonleap; }

    var daysCounter = 0;
    var numDays = 0;
    var monthDays = 0;

    if (targetDays < 0) {

        while (daysCounter < (targetDays * -1)) {

            if (daysCounter == 0) {
                if ((targetDays * -1) < thisDays) {
                    break;
                } else {
                    daysCounter = thisDays;
                }
            } else {
                numDays = monthArr[currMonth - 1];
                daysCounter += parseInt(numDays)
            }

            if (daysCounter > (targetDays * -1)) {
                break;
            }

            currMonth = currMonth - 1;

            if (currMonth == 0) {
                currYear = currYear - 1;
                if ((currYear % 4) == 0) {
                    if ((currYear % 100) == 0 && (currYear % 400) != 0) { monthArr = nonleap; }
                    else { monthArr = leap; }
                }
                else { monthArr = nonleap; }
                currMonth = 12;
            }
        }

        t = this.getTime();
        t += (targetDays * 86400000);
        this.setTime(t)
        var thisDate = new Date(currYear, currMonth - 1, this.getDate())
        return thisDate;

    } else {

        var diffDays = monthArr[currMonth - 1] - thisDays;

        numDays = 0;
        var startedC = true;

        while (daysCounter < targetDays) {

            if (daysCounter == 0 && startedC == true) {
                monthDays = thisDays;
                startedC = false;
            } else {
                monthDays++;
                daysCounter++;

                if (monthDays > monthArr[currMonth - 1]) {
                    currMonth = currMonth + 1;
                    monthDays = 1;
                }

            }

            if (daysCounter > targetDays) {
                break;
            }

            if (currMonth == 13) {
                currYear = currYear + 1;
                if ((currYear % 4) == 0) {
                    if ((currYear % 100) == 0 && (currYear % 400) != 0) { monthArr = nonleap; }
                    else { monthArr = leap; }
                }
                else { monthArr = nonleap; }
                currMonth = 1;
            }
        }

        var thisDate = new Date(currYear, currMonth - 1, monthDays)
        return thisDate;
    }
}