function makeGroups(memberList, numGroup, exceptList){

  for(let i = 0; i < exceptList.length ; i++){

    let target = memberList.indexOf(exceptList[i]);
    memberList.splice(target,1);
  
  }

  if(memberList.length < numGroup){

    return false;
  }


  let no = 0;
  const groups = [];
  for( let i = 0; i < numGroup; i++){

    //const newGroup = [memberList.length / numGroup + 1];
    //groups[i] = newGroup;

    groups.push([]);
  }

  let groupLimit = Math.floor(memberList.length / numGroup);
  for(let i = 0; i < memberList.length; i++){

    let rand = Math.random();
    rand = Math.floor(rand*numGroup);   

    console.log(rand)

    if(groups[rand].length < groupLimit){

      groups[rand].push(memberList[no++]);
    }
    else{

      if(groups[rand].length == groupLimit && memberList.length - no  <= 1 && no == memberList.length){

        groups[rand].push(memberList[no++]);
        continue;
      }

      i--;
      continue;
    }
  }

  return groups;
}

function main() {

  const member = ["a", "b", "c", "d", "e", "f", "g", "h", "i"];
  const except = ["a", "d", "f"];

 console.log(makeGroups(member, 6, except));
}
