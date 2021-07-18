import { Context } from "@azure/functions"

export class Logger {
        private static loggerType: LoggerType;
        private static context: Context;

        static Initialize(loggerType: LoggerType, context?: Context): void {
                this.loggerType = loggerType;
                this.context = context;
        }
        static log(object: any): void {
                if (this.loggerType === LoggerType.Console) {
                        console.log(object);
                }
                else if (this.loggerType === LoggerType.FunctionContext && this.context != undefined) {
                        this.context.log(object);
                }

        }
}
export enum LoggerType {
        Console = 'Console',
        FunctionContext = 'Function Context'
}