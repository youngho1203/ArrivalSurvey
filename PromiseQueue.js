/**
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
/**
 * @todo using promise queue : problem is that there is no event for cell value modification finish, so we don't know when ....
 */
function PromiseQueue() {
	this.queue = [];
  this.pendingPromise = false;
}

/**
 * enqueue
 */
PromiseQueue.prototype.enqueue = function(promise) {
  return new Promise((resolve, reject) => {
    this.queue.push({
      promise,
      resolve,
      reject,
    });
    this.dequeue();
  });  
}
  
/**
 * dequeue
 */
PromiseQueue.prototype.dequeue = function() {
  if (this.workingOnPromise) {
    return false;
  }
  const item = this.queue.shift();
  if (!item) {
    return false;
  }
  try {
    this.workingOnPromise = true;
    item.promise()
    .then((value) => {
      this.workingOnPromise = false;
      item.resolve(value);
      this.dequeue();
    })
    .catch(err => {
      this.workingOnPromise = false;
      item.reject(err);
      this.dequeue();
    })
  } catch (err) {
    this.workingOnPromise = false;
    item.reject(err);
    this.dequeue();
  }
  return true;
}