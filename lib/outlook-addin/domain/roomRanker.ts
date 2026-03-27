import type { RoomAvailability } from "../graph/schedule";

export interface RankingOptions {
  preferredCapacity?: number;
  preferredTags?: string[];
  preferAccessible?: boolean;
  preferVideoEnabled?: boolean;
}

/**
 * Rank rooms based on availability and preferences
 */
export function rankRooms(
  rooms: RoomAvailability[],
  options: RankingOptions = {}
): RoomAvailability[] {
  return [...rooms].sort((a, b) => {
    // First priority: availability
    if (a.isAvailable !== b.isAvailable) {
      return a.isAvailable ? -1 : 1;
    }

    // Calculate scores for each room
    const scoreA = calculateRoomScore(a, options);
    const scoreB = calculateRoomScore(b, options);

    // Higher score is better
    return scoreB - scoreA;
  });
}

/**
 * Calculate a preference score for a room
 */
function calculateRoomScore(
  roomAvail: RoomAvailability,
  options: RankingOptions
): number {
  const room = roomAvail.room;
  let score = 0;

  // Capacity fit (prefer closest to desired without being smaller)
  if (options.preferredCapacity) {
    const diff = room.capacity - options.preferredCapacity;
    if (diff >= 0) {
      // Room is big enough - prefer smaller rooms that still fit
      score += 10 - Math.min(diff, 10);
    } else {
      // Room is too small - penalize
      score -= 20;
    }
  } else {
    // No preference - slightly prefer smaller rooms for efficiency
    score += Math.max(0, 10 - room.capacity);
  }

  // Tag matching
  if (options.preferredTags && options.preferredTags.length > 0) {
    const roomTags = room.tags.map((t) => t.toLowerCase());
    for (const tag of options.preferredTags) {
      if (roomTags.includes(tag.toLowerCase())) {
        score += 5;
      }
    }
  }

  // Accessibility preference
  if (options.preferAccessible && room.isWheelChairAccessible) {
    score += 15;
  }

  // Video capability preference
  if (options.preferVideoEnabled && room.videoDeviceName) {
    score += 10;
  }

  return score;
}

/**
 * Group rooms by availability status
 */
export function groupRoomsByAvailability(rooms: RoomAvailability[]): {
  available: RoomAvailability[];
  unavailable: RoomAvailability[];
} {
  return {
    available: rooms.filter((r) => r.isAvailable),
    unavailable: rooms.filter((r) => !r.isAvailable),
  };
}

/**
 * Get room feature tags for display
 */
export function getRoomFeatures(room: RoomAvailability["room"]): string[] {
  const features: string[] = [];

  if (room.capacity > 0) {
    features.push(`${room.capacity} people`);
  }

  if (room.videoDeviceName) {
    features.push("Video");
  }

  if (room.audioDeviceName) {
    features.push("Audio");
  }

  if (room.displayDeviceName) {
    features.push("Display");
  }

  if (room.isWheelChairAccessible) {
    features.push("Accessible");
  }

  if (room.floorLabel) {
    features.push(`Floor ${room.floorLabel}`);
  }

  // Add custom tags
  for (const tag of room.tags) {
    if (!features.includes(tag)) {
      features.push(tag);
    }
  }

  return features;
}
