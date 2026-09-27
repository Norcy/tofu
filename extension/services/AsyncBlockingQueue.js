/**
 * Class AsyncBlockingQueue
 */
export default class AsyncBlockingQueue {
    constructor() {
        this.items = [];
        this.resolvers = [];
    }

    enqueue(item) {
        this.items.push(item);
        const resolve = this.resolvers.shift();
        if (resolve) resolve();
    }

    async dequeue() {
        if (!this.items.length) {
            await new Promise(resolve => this.resolvers.push(resolve));
        }
        return this.items.shift();
    }

    isEmpty() {
        return this.items.length === 0;
    }

    isBlocked() {
        return this.resolvers.length > 0;
    }

    clear() {
        this.items.length = 0;
    }

    get length() {
        return this.items.length;
    }
}
