// ----- Configure these parameters if needed -----

const cpaMembers = 9; // Number of CPA people
const dailySlots = 16; // Number of slots (30 min/slot) per day
const maxRandomRuns = 1000; // Max number of tries to randomly allocate Drop Ins and 1on1s.

// Constants

const eventInt = {
  FREE: 0,
  CLASS_OR_BUSY: 1,
  DROP_IN: 2,
  ONE_ON_ONE: 3,
  TEAM_MEETING: 4
};

const days = 5;
const totalSlots = dailySlots * days;
const book = SpreadsheetApp.getActiveSpreadsheet();
const resultTemplateSheet = book.getSheetByName("ResultTemplate");

//  ***** Run this function *****

function main() {
  console.log("Started analyzing...");

  const inputSheet = book.getSheetByName("Input");
  const inputRange = inputSheet.getRange(3, 2, cpaMembers, totalSlots);
  

  // Input class schedule
  const allScheduleArr = inputRange.getValues();

  // Decide team meeting time
  const slotIndexToStartTeamMeeting = findTeamMeetingSlots(allScheduleArr);

  if (!slotIndexToStartTeamMeeting) {
    console.log("There is no available slot for team meeting. Ask CPAs to set aside more time.");
    return;
  }

  allocateCPAEvents(slotIndexToStartTeamMeeting, allScheduleArr);
}


// ----- Utils -----

function transpose(array){
  return array[0].map((_, i) => array.map(row => row[i]));
}

function sum(array){
  return array.reduce((acc, num) => acc + num, 0);
}

function getPopularIndexArr(array) {
  return array
    .map((value, index) => ({ index, value }))
    .sort((a, b) => b.value - a.value)
    .map(pair => pair.index);
}

function isAllElementEqualToThreshold(array, threshold) {
  return array.every(element => element === threshold);
}

function pickRandomElements(arr, count) {
  const pickedElements = new Set();
  while (pickedElements.size < count) {
    const index = Math.floor(Math.random() * arr.length);
    pickedElements.add(arr[index]);
  }
  return Array.from(pickedElements);
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]]; // 配列の要素を交換
  }
  return array;
}


// ----- Team Meeting -----

function findTeamMeetingSlots(schedule){
  const slotIndexToStartTeamMeeting = [];
  const transposedSchedule = transpose(schedule); //80行9列

  for (let i=0; i<totalSlots-1; i++){
    if (i%dailySlots === dailySlots-1) continue; //日を跨ぐ場合

    const slot = transposedSchedule[i];
    const nextSlot = transposedSchedule[i+1];
    
    if (sum(slot) + sum(nextSlot) === 0){
      slotIndexToStartTeamMeeting.push(i);
    }
  }

  return slotIndexToStartTeamMeeting;
}


function allocateCPAEvents(meetingSlots, schedule) {
  const optionsCount = meetingSlots.length;

  for (let i = 0; i < optionsCount; i++) {
    console.log(`Calculating ${i + 1}/${optionsCount}...`);

    let updatedSchedule = schedule.map(row => [...row]);

    // Register team meeting
    const startIdx = meetingSlots[i];
    for (let j = 0; j < cpaMembers; j++) {
      updatedSchedule[j][startIdx] = eventInt.TEAM_MEETING;
      updatedSchedule[j][startIdx + 1] = eventInt.TEAM_MEETING;
    }

    // Schedule Drop-In
    const popularClassIndices = getPopularIndexArr(schedule);
    const dropInSlots = popularClassIndices.filter(index =>
      index !== startIdx 
      && index !== startIdx + 1 
      && index % dailySlots >= 4 
      && index % dailySlots <= 9
    );

    updatedSchedule = scheduleDropIn(updatedSchedule, dropInSlots);

    // Schedule 1on1
    updatedSchedule = schedule1on1(updatedSchedule);

    // Output to Result sheet
    outputToResultSheet(updatedSchedule, i)
  }
}


// ----- Drop In -----

function scheduleDropIn(schedule, dropInSlots) {
  let updatedSchedule;
  let numOfDropInArr = Array(cpaMembers).fill(0);

  // Ideally, every Drop In starts at 11:00 or 12:30 and lasts for 90 min.
  const idealDropInStartIndices = [4, 7, 20, 23, 36, 39, 52, 55, 68, 71]; 
  
  let randomRunCountForIdeal = 1;
  while (!isAllElementEqualToThreshold(numOfDropInArr, 6)) {
    if (randomRunCountForIdeal === 1){
      console.log("Looking for ideal Drop In schedule...");
    }

    if (randomRunCountForIdeal === maxRandomRuns){
      console.log("Unable to find ideal Drop In schedule.");
      break;
    }

    updatedSchedule = schedule.map(row => [...row]);
    numOfDropInArr = Array(cpaMembers).fill(0);

    for (let startIdx of idealDropInStartIndices) {
      const candidates = getCandidateIndices(schedule, numOfDropInArr, startIdx, 3);

      if (candidates.length >= 2) {
        const selected = pickRandomElements(candidates, 2);
        updateDropInSchedule(updatedSchedule, numOfDropInArr, selected, startIdx, 3);
      } else if (candidates.length === 1) {
        updateDropInSchedule(updatedSchedule, numOfDropInArr, candidates, startIdx, 3);
      }
    }

    randomRunCountForIdeal += 1;
  }

  let randomRunCount = 1;
  while (!isAllElementEqualToThreshold(numOfDropInArr, 6)) {
    if (randomRunCount === 1){
      console.log("Started random allocation for Drop In...");
    }

    updatedSchedule = schedule.map(row => [...row]);
    numOfDropInArr = Array(cpaMembers).fill(0);

    if (randomRunCount === maxRandomRuns){
      throw new Error("Unable to allocate Drop In randomly. Try running again. You may need to increase maxRandomRuns value.")
    }

    for (let slotIdx of dropInSlots) {
      const candidates = getCandidateIndices(schedule, numOfDropInArr, slotIdx, 1);

      if (candidates.length >= 2) {
        const selected = pickRandomElements(candidates, 2);
        updateDropInSchedule(updatedSchedule, numOfDropInArr, selected, slotIdx, 1);
      } else if (candidates.length === 1) {
        updateDropInSchedule(updatedSchedule, numOfDropInArr, candidates, slotIdx, 1);
      }
    }

    randomRunCount += 1;
  }

  return updatedSchedule;
}

function getCandidateIndices(schedule, dropInArr, startIndex, duration) {
  return Array.from({ length: cpaMembers }, (_, i) => i)
    .filter(i => 
      Array.from({ length: duration }, (_, j) => schedule[i][startIndex + j])
        .every(val => val === eventInt.FREE) 
      &&
      dropInArr[i] < 6
    )
  ;
}

function updateDropInSchedule(schedule, dropInArr, indices, startIndex, duration) {
  indices.forEach(i => {
    for (let j = 0; j < duration; j++) {
      schedule[i][startIndex + j] = eventInt.DROP_IN;
      dropInArr[i]++;
    }
  });
}


// ----- 1on1 -----

function schedule1on1(schedule) {
  // Try to allocate 2h*2 1on1 for every CPA.

  const busyPersonIndices = getPopularIndexArr(
    schedule.map(personSchedule => 
      personSchedule.filter(slotValue => slotValue!==eventInt.FREE).length
    )
  );

  let numOf1on1ForPersonArr = Array(cpaMembers).fill(0);
  let numOf1on1ForSlotArr = Array(totalSlots).fill(0);
  let numOf1on1ForDayArr = Array(days).fill(0);

  let updatedSchedule = schedule.map(row => [...row]);

  // From the busiest person
  console.log("Looking for ideal 1on1 schedule...");

  for (personIndex of busyPersonIndices){
    for (let i = 0; i < totalSlots; i++) {
      const personSchedule = updatedSchedule[personIndex];
      const dayIndex = Math.floor(i / dailySlots);
      const mod16 = i % dailySlots;

      if (
        (mod16 >= 0 && mod16 <= 12)
        && numOf1on1ForDayArr[dayIndex] < 16 
        && numOf1on1ForPersonArr[personIndex] < 8
        && Array.from({length: 4}, (_, index) => personSchedule[i+index]).every(value => value === eventInt.FREE)
        && Array.from({length: 4}, (_, index) => numOf1on1ForSlotArr[i+index]).every(value => value < 2)
      ){
        for (let j = 0; j < 4; j++) {
          personSchedule[i + j] = eventInt.ONE_ON_ONE;
          numOf1on1ForSlotArr[i + j] += 1;
        }
        numOf1on1ForDayArr[dayIndex] += 4;
        numOf1on1ForPersonArr[personIndex] += 4;
        i = dayIndex * dailySlots + (dailySlots - 1);
      }
    }
  };

  // Allocate 1on1 randomly
  let randomRunCount = 1;
  while(!isAllElementEqualToThreshold(numOf1on1ForPersonArr, 8)){
    if (randomRunCount === 1){
      console.log("Started random allocation for 1on1 schedule...");
    }

    if (randomRunCount === maxRandomRuns){
      throw new Error("Unable to allocate Drop In randomly. Try running again. You may need to increase maxRandomRuns value.")
    }

    numOf1on1ForPersonArr = Array(cpaMembers).fill(0);
    numOf1on1ForSlotArr = Array(totalSlots).fill(0);
    numOf1on1ForDayArr = Array(days).fill(0);

    updatedSchedule = schedule.map(row => [...row]);

    const peopleIndex = shuffleArray(updatedSchedule);
    for (personIndex of peopleIndex){
      for (let i = 0; i < totalSlots; i++){
        const personSchedule = updatedSchedule[personIndex];
        const dayIndex = Math.floor(i / dailySlots);
        const mod16 = i % dailySlots;

        if (
          (mod16 >= 0 && mod16 <= 12)
          && numOf1on1ForDayArr[dayIndex] < 16 
          && numOf1on1ForPersonArr[personIndex] < 8
          && Array.from({length: 4}, (_, index) => personSchedule[i+index]).every(value => value === eventInt.FREE)
          && Array.from({length: 4}, (_, index) => numOf1on1ForSlotArr[i+index]).every(value => value < 2)
        ){
          for (let j = 0; j < 4; j++) {
            personSchedule[i + j] = eventInt.ONE_ON_ONE;
            numOf1on1ForSlotArr[i + j]++;
          }
          numOf1on1ForDayArr[dayIndex] += 4;
          numOf1on1ForPersonArr[personIndex] += 4;
          i = dayIndex * dailySlots + (dailySlots - 1);
        }
      }
    }

    randomRunCount += 1;
  }

  return updatedSchedule;
}


function outputToResultSheet(schedule, index){
  const resultSheetName = `Result_${index + 1}`;
  const existingSheet = book.getSheetByName(resultSheetName);

  if (existingSheet){
    book.deleteSheet(existingSheet);
  }

  const resultSheet = resultTemplateSheet.copyTo(book).setName(resultSheetName);
  resultSheet.getRange(3, 2, cpaMembers, totalSlots).setValues(schedule);
}
