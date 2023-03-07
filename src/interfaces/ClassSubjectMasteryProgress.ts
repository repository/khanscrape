export interface ClassSubjectMasteryProgress {
  data: {
    classroom: Classroom;
    topicById: TopicByID;
  };
}

export interface Classroom {
  __typename: "StudentList";
  cacheId: string;
  id: string;
  students: Student[];
}

export interface Student {
  __typename: "Student";
  coachNickname: string;
  id: string;
  kaid: string;
  subjectProgress: SubjectProgress;
}

export interface SubjectProgress {
  __typename: "SubjectProgress";
  currentMastery: CurrentMastery;
  unitProgresses: UnitProgress[];
}

export interface CurrentMastery {
  __typename: "CurationItemMastery";
  percentage: number;
  pointsAvailable: number;
  pointsEarned: number;
}

export interface UnitProgress {
  __typename: "UnitProgress";
  currentMastery: CurrentMastery;
  topic: Topic;
}

export interface Topic {
  __typename: "Topic";
  id: string;
}

export interface TopicByID {
  __typename: "Topic";
  childTopics: ChildTopic[];
  id: string;
}

export interface ChildTopic {
  __typename: "Topic";
  id: string;
  translatedTitle: string;
}
