#include <iostream>
#include <vector>
using namespace std;

#include <stdlib.h>
#include <math.h>

// How many "ranges".
const int NUMRANGES = 50;

// How many entries are in a "range".
const int NUMENTRIES = 50;

// How many samples do we take of a range.
const int Samples = 30;

// The data ranges.
vector<float> data[NUMRANGES];

// Their associated impact.
vector<float> impact[NUMRANGES];

// Add up a vector.
template <class CLASS>
CLASS sumFunc (vector<CLASS> arr, int N) {
  CLASS sum = 0;
  for (int i = 0; i < N; i++) {
    sum += arr[i];
  }
  return sum;
#if 0
  if ((long) sum % 2 == 0) {
    return 1;
  } else {
    return 0;
  }
  //return sum / (CLASS) N;
#endif
}

// Recalculate the "spreadsheet".
template <class CLASS>
CLASS recalc() {
  CLASS sum = 0;
  for (int i = 0; i < NUMRANGES; i++) {
    sum += sumFunc<CLASS>(data[i], NUMENTRIES);
  }
  return sum / NUMRANGES;
}


// For generating exponential distributions.
float exponential (float lambda) {
  return -1.0 / lambda * log(1.0 - drand48());
}

int
main()
{

  // Initialize the vectors.
  for (int i = 0; i < NUMRANGES; i++) {
    data[i].resize (NUMENTRIES);
    impact[i].resize (NUMENTRIES);
  }

  // Fill the vectors with random numbers (exponentially distributed).
  for (int j = 0; j < NUMRANGES; j++) {
    for (int i = 0; i < NUMENTRIES; i++) {
      //      data[j][i] = (float) (lrand48() % 100); // exponential (4.0);
      data[j][i] = exponential (4.0);
      impact[j][i] = 0.0;
      //      cout << data[j][i] << endl;
    }
    //    cout << "..." << endl;
  }
  //  cout << endl;
  //  cout << "----" << endl;

  // Compute the probability that we will swap any given item.
  double prob = (double) Samples / (double) NUMENTRIES;

  // Compute the baseline calculation.
  float actualSum = recalc<float>();

  // Bootstrapping.

#if 0
  for (int k = 0; k < NUMRANGES; k++) {
    for (int i = 0; i < NUMENTRIES; i++) {
      float sum = 0;
      vector<float> boot (NUMENTRIES);
      // Generate a number of bootstrapped distributions.
      for (int j = 0; j < 10 * Samples; j++) {
	for (int q = 0; q < NUMENTRIES; q++) {
	  int index;
	  // Find a random item for the bootstrap, but not the current one,
	  // which we are excluding.
	  do {
	    index = lrand48() % NUMENTRIES;
	  } while (index == i);
	  boot[q] = data[k][index];
	}
      }
      vector<float> backup (NUMENTRIES);
      backup = data[k];
      data[k] = boot;
      sum += fabs(recalc<float>() - actualSum);
      data[k] = backup;
      impact[k][i] = sum / (10 * Samples);
    }
  }

#else
  // Perform the "perturbation."
  for (int k = 0; k < NUMRANGES; k++) {
    for (int i = 0; i < NUMENTRIES; i++) {
      float sum = 0;
      int perturbationCount = 0;
      for (int j = 0; j < NUMENTRIES; j++) {
	double r = drand48();
	if (r <= prob) {
	  int swapIndex;
	  // Find a random item to swap but don't swap with itself.
	  do {
	    swapIndex = lrand48() % NUMENTRIES;
	  } while (swapIndex == i);
	  float orig = data[k][i];
	  data[k][i] = data[k][swapIndex];
	  sum += fabs(recalc<float>() - actualSum);
	  perturbationCount++;
	  data[k][i] = orig;
	}
      }
      //      cout << "[" << perturbationCount << "]" << endl;
      impact[k][i] = sum / perturbationCount;
    }
  }
#endif

  // Normalize the impacts by converting them into z-scores.
  // First, compute the means.
  float mean[NUMRANGES];
  for (int k = 0; k < NUMRANGES; k++) {
    mean[k] = 0;
    for (int i = 1; i < NUMENTRIES; i++) {
      mean[k] += impact[k][i];
    }
    mean[k] /= NUMENTRIES;
  }

  // Now, the standard deviations.
  float stddev[NUMRANGES];
  for (int k = 0; k < NUMRANGES; k++) {
    stddev[k] = 0;
    for (int i = 1; i < NUMENTRIES; i++) {
      float diff = impact[k][i] - mean[k];
      stddev[k] += diff * diff;
    }
    stddev[k] /= NUMENTRIES;
    stddev[k] = sqrt(stddev[k]);
  }

  // Now normalize.
  for (int k = 0; k < NUMRANGES; k++) {
    for (int i = 1; i < NUMENTRIES; i++) {
      impact[k][i] = fabs(impact[k][i] - mean[k]) / stddev[k];
    }
  }
  
  for (int k = 0; k < NUMRANGES; k++) {
    float v = 0;
    for (int i = 0; i < NUMENTRIES; i++) {
      v += impact[k][i];
    }
    cout << v / NUMENTRIES << endl;
    //    cout << "..." << endl;
    //    cout << endl;
  }

#if 0
  // Generate an Erlang distribution.
  vector<float> erl (NUMENTRIES);
  for (int i = 0; i < NUMENTRIES; i++) {
    const int R = 2;
    const float lambda = 0.5;
    erl[i] = 0;
    for (int j = 0; j < R; j++) {
      erl[i] += exponential (lambda);
    }
  }
  // Normalize.
  float min = erl[0];
  float max = erl[0];
  for (int i = 0; i < NUMENTRIES; i++) {
    if (erl[i] < min) {
      min = erl[i];
    }
    if (erl[i] > max) {
      max = erl[i];
    }
  }
  for (int i = 0; i < NUMENTRIES; i++) {
    erl[i] = (erl[i] - min) / (max - min);
    cout << erl[i] << endl;
  }
#endif

}
