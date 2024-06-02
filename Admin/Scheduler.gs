/*
 * Usage:
 *  const schedulable_procedure = new Scheduler(
 *      function schedulable_procedure() { // the name must coincide!
 *          // do things
 *      },
 *      "schedulable_procedure.schedule", // for property storage
 *      function schedule_generator() {
 *          let date = new Date();
 *          date.setMinutes(date.getMinutes() + 5);
 *          return [{date, args: []}];
 *      }
 *  );
 */

class Scheduler {

    /**
     * @param {Function} target
     * @param {string} schedule_key
     * @param {() => Array<{date: Date, args: any[]}>} schedule_generator
     */
    constructor(target, schedule_key, schedule_generator) {
        this.target = target;
        this.schedule_key = schedule_key;
        this.schedule_generator = schedule_generator;
    }

    today() {
        let now = new Date();
        let schedule = this.schedule_generator()
            .filter(({date}) => date.valueOf() > now.valueOf());
        this.store_schedule(schedule);
        this.setup_next(schedule);
    }

    /**
     * @param {Array<{date: Date, args: any[]}>} schedule
     */
    store_schedule(schedule) {
        if (schedule.length == 0) {
            PropertiesService.getDocumentProperties()
                .deleteProperty(this.schedule_key);
            return;
        }
        let schedule_encoded = JSON.stringify(schedule.map(
            ({date, args}) => [date.valueOf(), args]
        ));
        PropertiesService.getDocumentProperties()
            .setProperty(this.schedule_key, schedule_encoded);
    }

    /**
     * @return {Array<{date: Date, args: any[]}>} args
     */
    retrieve_schedule() {
        let schedule_encoded = PropertiesService.getDocumentProperties()
            .getProperty(this.schedule_key);
        if (schedule_encoded == null) {
            return [];
        }
        return JSON.parse(schedule_encoded).map(
            ([timestamp, args]) => ({date: new Date(timestamp), args})
        );
    }

    /**
     * @param {Array<{date: Date, args: any[]}>} schedule
     */
    setup_next(schedule) {
        this.not_today();
        if (schedule.length == 0) {
            return;
        }
        ScriptApp.newTrigger(this.target.name + ".run")
            .timeBased().at(schedule[0].date)
            .create();
    }

    run() {
        let schedule = this.retrieve_schedule();
        let next = schedule.shift();
        let error = null;
        try {
            this.setup_next(schedule);
            this.store_schedule(schedule);
        } catch (err) {
            err = error;
            console.log(error);
        }
        if (next == undefined) {
            throw new Error("no arguments to process");
        }
        let {args} = next;
        if (args.length == 0) {
            console.log(
                "running " + this.target.name + " " );
        } else {
            console.log(
                "running " + this.target.name + " " +
                "with args " + JSON.stringify(args) );
        }
        (this.target)(...args);
        if (error != null) {
            throw error;
        }
    }

    not_today() {
        for (let trigger of ScriptApp.getProjectTriggers()) {
            if (trigger.getHandlerFunction().startsWith(
                this.target.name + ".run" ))
            {
                ScriptApp.deleteTrigger(trigger);
            }
        }
    }

    never() {
        for (let trigger of ScriptApp.getProjectTriggers()) {
            if (trigger.getHandlerFunction().startsWith(this.target.name))
                ScriptApp.deleteTrigger(trigger);
        }
    }

}

