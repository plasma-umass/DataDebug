#include <iostream>
#include <vector>
using namespace std;

#include <stdlib.h>
#include <math.h>

// How many "ranges".
const int NUM = 50;

// How many entries are in a "range".
const int N = 100;

// How many samples do we take of a range.
const int Samples = 30;

// The data ranges.
vector<float> data[NUM];

// Their associated impact.
vector<float> impact[NUM];

// Add up a vector.
template <class CLASS>
CLASS sumFunc (vector<CLASS> arr, int N) {
  CLASS sum = 0;
  for (int i = 0; i < N; i++) {
    sum += arr[i];
  }
  return sum;
}

// Recalculate the "spreadsheet".
template <class CLASS>
CLASS recalc() {
  CLASS sum = 0;
  for (int i = 0; i < NUM; i++) {
    sum += sumFunc<CLASS>(data[i], N);
  }
  return sum;
}


// For generating exponential distributions.
float exponential (float lambda) {
  return -1.0 / lambda * log(1.0 - drand48());
}

int
main()
{

  // Initialize the vectors.
  for (int i = 0; i < NUM; i++) {
    data[i].resize (N);
    impact[i].resize (N);
  }

  // Fill the vectors with random numbers (exponentially distributed).
  for (int j = 0; j < NUM; j++) {
    for (int i = 0; i < N; i++) {
      data[j][i] = exponential (4.0);
      impact[j][i] = 0.0;
      cout << data[j][i] << endl;
    }
    cout << "..." << endl;
  }
  cout << endl;
  cout << "----" << endl;

  // Compute the probability that we will swap any given item.
  double prob = (double) Samples / (double) N;

  // Compute the baseline calculation.
  float actualSum = recalc<float>();

  // Perform the "perturbation."
  for (int k = 0; k < NUM; k++) {
    for (int i = 0; i < N; i++) {
      float sum = 0;
      int perturbationCount = 0;
      for (int j = 0; j < N; j++) {
	double r = drand48();
	if (r <= prob) {
	  int swapIndex;
	  // Find a random item to swap but don't swap with itself.
	  do {
	    swapIndex = lrand48() % N;
	  } while (swapIndex == i);
	  float orig = data[k][i];
	  data[k][i] = data[k][swapIndex];
	  sum += fabs(recalc<float>() - actualSum);
	  perturbationCount++;
	  data[k][i] = orig;
	}
      }
      cout << "[" << perturbationCount << "]" << endl;
      impact[k][i] = sum / perturbationCount;
    }
  }

  for (int k = 0; k < NUM; k++) {
    float minImpact = impact[k][0];
    float maxImpact = impact[k][0];
    for (int i = 1; i < N; i++) {
      if (impact[k][i] < minImpact) {
	minImpact = impact[k][i];
      }
      if (impact[k][i] > maxImpact) {
	maxImpact = impact[k][i];
      }
    }
    for (int i = 0; i < N; i++) {
      impact[k][i] = (impact[k][i] - minImpact) / (maxImpact - minImpact);
    }
  }
  
  for (int k = 0; k < NUM; k++) {
    for (int i = 0; i < N; i++) {
      cout << impact[k][i] << endl;
    }
    cout << "..." << endl;
    cout << endl;
  }

  // Generate an Erlang distribution.
  vector<float> erl (N);
  for (int i = 0; i < N; i++) {
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
  for (int i = 0; i < N; i++) {
    if (erl[i] < min) {
      min = erl[i];
    }
    if (erl[i] > max) {
      max = erl[i];
    }
  }
  for (int i = 0; i < N; i++) {
    erl[i] = (erl[i] - min) / (max - min);
    cout << erl[i] << endl;
  }
}
