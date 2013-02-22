// -*- C++ -*-

#ifndef SHUFFLE_H_
#define SHUFFLE_H_

#include <stdlib.h>
#include <vector>
using namespace std;

template <class TYPE>
class fyshuffle {
public:

  static void inplace (vector<TYPE>& vec)
  {
    for (auto i = vec.size()-1; i > 0; i--) {
      auto j = lrand48() % (i+1);
      swap (vec[i], vec[j]);
    }
  }

  static void transform (const vector<TYPE>& in,
			 vector<TYPE>& out)
  {
    out[0] = in[0];
    for (auto i = 1; i < in.size(); i++) {
      auto j = lrand48() % (i+1);
      out[i] = out[j];
      out[j] = in[i];
    }
  }

};


#endif
