export class Birthday {
    Title: string;
    Day: number;
    Month: number;

    // Kas sünnipäev on täna
    public isToday() : boolean {
        const currentDate = new Date();
        const currentMonth = currentDate.getMonth() + 1;
        const currentDay = currentDate.getDate();

        return (this.Day === currentDay &&
                this.Month === currentMonth);
    }

    // Kas sünnipäev on käesoleval kuul
    public isAtCurrentMonth() : boolean {
        // new Date() tagastab käesoleva ajahetke
        // getMonth() tagastab kuupäeva küljest kuu, kuid kuude
        // loendamine JavaScriptis algab nullist, seega lisame ühe juurde
        const currentMonth = new Date().getMonth() + 1;

        // Kui käesolev kuu ja sünnipäeva kuu on võrdsed, siis on 
        // sünnipäev käesoleval kuul
        return (this.Month === currentMonth);
    }

    // Koosta kuupäev kujul "pp.kk" - näiteks "01.06"
    public formatAsDayAndMonth() : string {
        let day = this.Day.toString();
        if(day.length === 1) {
            day = "0" + day;
        }

        let month = this.Month.toString();
        if(month.length === 1) {
            month = "0" + month;
        }

        return day + "." + month;
    }

    public static sort(birthdays: Birthday[]) : Birthday[] {
        return birthdays.sort(Birthday._compareBirthdays);
    }

    private static _compareBirthdays(birthday1: Birthday, birthday2: Birthday) : number {
        let returnValue = 0;

        if(birthday1.Month < birthday2.Month) {
            return -1;
        }
        if(birthday2.Month < birthday1.Month) {
            return 1;
        }

        if(birthday1.Day < birthday2.Day) {
            returnValue = -1;
        }
        if(birthday2.Day < birthday1.Day) {
            returnValue = 1;
        }

        return returnValue;
    }

    public static fromRandomObject(obj: Birthday) : Birthday {
        const birthday = new Birthday();
        birthday.Title = obj.Title;
        birthday.Day = parseInt(obj.Day.toString());
        birthday.Month = parseInt(obj.Month.toString());

        return birthday;
    }

    public static fromRandomObjects(obj: Birthday[]) : Birthday[] {
        const birthdays = new Array<Birthday>();

        obj.forEach(fakeBirthday => {
            birthdays.push(Birthday.fromRandomObject(fakeBirthday));
        });

        return birthdays;
    }
}